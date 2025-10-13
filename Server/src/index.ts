// src/index.ts — XQLite GraphQL server (single file)
// deps: npm i graphql graphql-yoga graphql-type-json better-sqlite3
// env:  PORT, XQL_DB, XQL_DEFAULT_KEY

import { createServer } from "http";
import { createYoga, createPubSub } from "graphql-yoga";
import {
  GraphQLObjectType,
  GraphQLInputObjectType,
  GraphQLSchema,
  GraphQLString,
  GraphQLList,
  GraphQLInt,
  GraphQLBoolean,
  GraphQLNonNull,
  GraphQLScalarType,
} from "graphql";
import GraphQLJSON from "graphql-type-json";
import Database from "better-sqlite3";
import fs from "fs";

// ========================= DB bootstrap =========================
const DB_PATH = process.env.XQL_DB ?? "./xqlite.db";
const db = new Database(DB_PATH);
db.pragma("journal_mode = WAL");
db.pragma("synchronous = NORMAL");
db.pragma("busy_timeout = 5000");

db.exec(`
CREATE TABLE IF NOT EXISTS _meta (
  k TEXT PRIMARY KEY,
  v TEXT
);
CREATE TABLE IF NOT EXISTS _presence (
  nickname TEXT PRIMARY KEY,
  sheet TEXT,
  cell TEXT,
  updated_at INTEGER
);
CREATE TABLE IF NOT EXISTS _locks (
  cell TEXT PRIMARY KEY,
  by TEXT NOT NULL,
  ts INTEGER NOT NULL
);
CREATE TABLE IF NOT EXISTS _audit (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  ts INTEGER NOT NULL,
  user TEXT,
  table_name TEXT,
  row_key TEXT,
  column_name TEXT,
  old_value TEXT,
  new_value TEXT,
  row_version INTEGER
);
CREATE TABLE IF NOT EXISTS _schema (
  table_name TEXT PRIMARY KEY,
  key_column TEXT NOT NULL
);
`);

db.exec(`INSERT OR IGNORE INTO _meta(k,v) VALUES('max_row_version','0')`);

type RowV = { v: number };
const getMaxRowVersion = db.prepare<[], RowV>(
  `SELECT CAST(v AS INTEGER) AS v FROM _meta WHERE k='max_row_version'`
);
const setMaxRowVersion = db.prepare<[number]>(
  `UPDATE _meta SET v=? WHERE k='max_row_version'`
);

function nextRowVersion(): number {
  const cur = Number(getMaxRowVersion.get()?.v ?? 0) + 1;
  setMaxRowVersion.run(cur);
  return cur;
}

// ========================= helpers =========================
function keepValueForStore(v: any) {
  if (v === null || v === undefined) return null;
  if (typeof v === "object") return JSON.stringify(v);
  if (v === true) return 1;
  if (v === false) return 0;
  return v;
}

function reviveValue(v: any) {
  if (typeof v === "string" && v.length > 0 && (v[0] === "{" || v[0] === "[")) {
    try { return JSON.parse(v); } catch { /* keep string */ }
  }
  return v;
}

function rowToCells(row: any, keyCol: string): Record<string, any> {
  const cells: Record<string, any> = {};
  for (const k of Object.keys(row)) {
    if (k === keyCol || k === "row_version" || k === "updated_at" || k === "deleted") continue;
    cells[k] = reviveValue(row[k]);
  }
  return cells;
}

function ensureTable(table: string, key: string) {
  const rec = db.prepare<[string], { key_column: string }>(
    `SELECT key_column FROM _schema WHERE table_name=?`
  ).get(table);
  if (rec?.key_column && rec.key_column !== key) {
    throw new Error(`Key mismatch for table '${table}': existing=${rec.key_column}, requested=${key}`);
  }
  db.exec(`
    CREATE TABLE IF NOT EXISTS "${table}" (
      "${key}" TEXT PRIMARY KEY,
      row_version INTEGER NOT NULL,
      updated_at INTEGER NOT NULL,
      deleted INTEGER NOT NULL DEFAULT 0
    )
  `);
  db.prepare(`INSERT OR IGNORE INTO _schema(table_name,key_column) VALUES(?,?)`).run(table, key);
}

type ColumnDefInput = { name: string; kind?: string; notNull?: boolean; check?: string };

const GqlColumnInfo = new GraphQLObjectType({
  name: "ColumnInfo",
  fields: {
    name: { type: new GraphQLNonNull(GraphQLString) },
    type: { type: GraphQLString },        // declared type (affinity용)
    notnull: { type: new GraphQLNonNull(GraphQLBoolean) },
    pk: { type: new GraphQLNonNull(GraphQLBoolean) },
  }
});


function addColumns(table: string, defs: ColumnDefInput[]) {
  const pragmaCols = db.prepare(`PRAGMA table_info("${table}")`).all() as any[];
  const have = new Set<string>(pragmaCols.map(c => String(c.name)));
  const colSqls: string[] = [];
  for (const d of defs) {
    if (!have.has(d.name)) {
      let typ = "TEXT";
      switch ((d.kind ?? "text").toLowerCase()) {
        case "int":
        case "integer": typ = "INTEGER"; break;
        case "real":
        case "float":
        case "double": typ = "REAL"; break;
        case "bool":
        case "boolean": typ = "INTEGER"; break; // 0/1
        case "date": typ = "INTEGER"; break;    // epoch ms
        case "json": typ = "TEXT"; break;       // JSON1 + TEXT
        default: typ = "TEXT";
      }
      const notNull = d.notNull ? " NOT NULL" : "";
      const check = d.check ? ` CHECK(${d.check})` : "";
      colSqls.push(`ALTER TABLE "${table}" ADD COLUMN "${d.name}" ${typ}${notNull}${check}`);
    }
  }
  const tx = db.transaction(() => colSqls.forEach(sql => db.exec(sql)));
  if (colSqls.length) tx();
}

function getSchemaRec(table: string): { key_column: string } | undefined {
  return db.prepare<[string], { key_column: string }>(
    `SELECT key_column FROM _schema WHERE table_name=?`
  ).get(table);
}
function getKeyCol(table: string): string {
  const r = getSchemaRec(table);
  if (!r) throw new Error(`Table not registered: ${table}`);
  return String(r.key_column);
}

const DEFAULT_KEY = process.env.XQL_DEFAULT_KEY ?? "id";
function ensureTableIfMissing(table: string, keyColHint?: string): string {
  const rec = getSchemaRec(table);
  if (rec?.key_column) return rec.key_column;
  const key = keyColHint ?? DEFAULT_KEY;
  ensureTable(table, key);
  return key;
}

function inferKind(v: any): "integer" | "real" | "boolean" | "json" | "text" {
  if (v === null || v === undefined) return "text";
  if (typeof v === "number") return Number.isInteger(v) ? "integer" : "real";
  if (typeof v === "boolean") return "boolean";
  if (typeof v === "object") return "json";
  return "text";
}

function ensureColumnsForSamples(table: string, sample: Record<string, any>, keyCol: string) {
  const reserved = new Set([keyCol, "row_version", "updated_at", "deleted"]);
  const defs: ColumnDefInput[] = [];
  for (const [name, val] of Object.entries(sample)) {
    if (reserved.has(name)) continue;
    defs.push({ name, kind: inferKind(val) });
  }
  if (defs.length) addColumns(table, defs);
}

function readSince(table: string, since: number) {
  const key = getKeyCol(table);
  const rows = db.prepare<[number], any>(`SELECT * FROM "${table}" WHERE row_version > ?`).all(since);
  return rows.map((r: any) => ({
    table,
    row_key: String(r[key]),
    row_version: r.row_version,
    deleted: !!r.deleted,
    cells: rowToCells(r, key),
  }));
}

// ========================= GraphQL types =========================
type CellEditInput = { table: string; row_key: string; column: string; value?: any };


const GqlPatch = new GraphQLObjectType({
  name: "RowPatch",
  fields: {
    table: { type: new GraphQLNonNull(GraphQLString) },
    row_key: { type: new GraphQLNonNull(GraphQLString) },
    row_version: { type: new GraphQLNonNull(GraphQLInt) },
    deleted: { type: new GraphQLNonNull(GraphQLBoolean) },
    cells: { type: GraphQLJSON },
  },
});

const GqlUpsertResult = new GraphQLObjectType({
  name: "UpsertResult",
  fields: {
    max_row_version: { type: new GraphQLNonNull(GraphQLInt) },
    errors: { type: new GraphQLList(GraphQLString) },
    conflicts: {
      type: new GraphQLList(new GraphQLObjectType({
        name: "Conflict",
        fields: {
          table: { type: GraphQLString },
          row_key: { type: GraphQLString },
          column: { type: GraphQLString },
          message: { type: GraphQLString },
          server_version: { type: GraphQLInt },
          local_version: { type: GraphQLInt },
          sheet: { type: GraphQLString },
          address: { type: GraphQLString },
          type: { type: GraphQLString },
        }
      }))
    }
  }
});

const GqlRowsResult = new GraphQLObjectType({
  name: "RowsResult",
  fields: {
    max_row_version: { type: new GraphQLNonNull(GraphQLInt) },
    patches: { type: new GraphQLList(GqlPatch) },
  }
});

const GqlOk = new GraphQLObjectType({
  name: "Ok",
  fields: { ok: { type: new GraphQLNonNull(GraphQLBoolean) } }
});

const GqlPresence = new GraphQLObjectType({
  name: "Presence",
  fields: {
    nickname: { type: new GraphQLNonNull(GraphQLString) },
    sheet: { type: GraphQLString },
    cell: { type: GraphQLString },
    updated_at: { type: new GraphQLNonNull(GraphQLInt) },
  }
});

const GqlAudit = new GraphQLObjectType({
  name: "Audit",
  fields: {
    ts: { type: new GraphQLNonNull(GraphQLInt) },
    user: { type: GraphQLString },
    table: { type: GraphQLString },
    row_key: { type: GraphQLString },
    column: { type: GraphQLString },
    old_value: { type: GraphQLString },
    new_value: { type: GraphQLString },
    row_version: { type: GraphQLInt },
  }
});

// Input types
const CellEditInputType = new GraphQLInputObjectType({
  name: "CellEditInput",
  fields: {
    table: { type: new GraphQLNonNull(GraphQLString) },
    row_key: { type: new GraphQLNonNull(GraphQLString) },
    column: { type: new GraphQLNonNull(GraphQLString) },
    value: { type: GraphQLJSON },
  }
});

const ColumnDefInputType = new GraphQLInputObjectType({
  name: "ColumnDefInput",
  fields: {
    name: { type: new GraphQLNonNull(GraphQLString) },
    kind: { type: GraphQLString },
    notNull: { type: GraphQLBoolean },
    check: { type: GraphQLString },
  }
});

// ========================= PubSub =========================
type ServerEventPayload = { max_row_version: number; patches: any[] };
const pubsub = createPubSub<{ events: [ServerEventPayload] }>();

// ========================= Mutations/Queries =========================
function upsertCells(cells: CellEditInput[]) {
  const errors: string[] = [];
  const byTable = new Map<string, Map<string, Record<string, any>>>();

  // group by (table, row_key)
  for (const c of cells) {
    if (!byTable.has(c.table)) byTable.set(c.table, new Map());
    const map = byTable.get(c.table)!;
    const key = String(c.row_key);
    const row = map.get(key) ?? {};
    row[c.column] = c.value;
    map.set(key, row);
  }

  const patches: any[] = [];
  const tx = db.transaction(() => {
    for (const [table, rowsMap] of byTable) {
      // auto-create table + columns
      const keyCol = ensureTableIfMissing(table, DEFAULT_KEY);
      const sample: Record<string, any> = {};
      for (const values of rowsMap.values())
        for (const [col, val] of Object.entries(values))
          if (!(col in sample)) sample[col] = val;
      ensureColumnsForSamples(table, sample, keyCol);

      // upsert
      for (const [rowKey, values] of rowsMap) {
        const rv = nextRowVersion();
        const now = Date.now();

        const old = db.prepare(`SELECT * FROM "${table}" WHERE "${keyCol}"=?`).get(rowKey) as any;

        db.prepare(
          `INSERT INTO "${table}"("${keyCol}",row_version,updated_at,deleted) VALUES (?,?,?,0)
           ON CONFLICT("${keyCol}") DO NOTHING`
        ).run(rowKey, rv, now);

        const cols = Object.keys(values);
        if (cols.length) {
          const sets = cols.map(c => `"${c}"=?`).join(", ");
          const vals = cols.map(c => keepValueForStore((values as any)[c]));
          db.prepare(
            `UPDATE "${table}" SET ${sets}, row_version=?, updated_at=?, deleted=0 WHERE "${keyCol}"=?`
          ).run(...vals, rv, now, rowKey);

          for (const c of cols) {
            const oldv = old ? old[c] : null;
            const newv = (values as any)[c];
            db.prepare(`INSERT INTO _audit(ts,user,table_name,row_key,column_name,old_value,new_value,row_version)
                        VALUES(?,?,?,?,?,?,?,?)`)
              .run(now, null, table, rowKey, c,
                oldv !== undefined ? String(oldv) : null,
                keepValueForStore(newv) !== null ? String(keepValueForStore(newv)) : null,
                rv);
          }
        }

        const rowFull = db.prepare(`SELECT * FROM "${table}" WHERE "${keyCol}"=?`).get(rowKey) as any;
        patches.push({
          table,
          row_key: String(rowKey),
          row_version: rowFull.row_version,
          deleted: !!rowFull.deleted,
          cells: rowToCells(rowFull, keyCol),
        });
      }
    }
  });

  try { tx(); }
  catch (e: unknown) { errors.push(e instanceof Error ? e.message : String(e)); }

  if (patches.length) {
    const r = getMaxRowVersion.get();
    pubsub.publish("events", { max_row_version: Number(r?.v ?? 0), patches });
  }

  return {
    max_row_version: Number(getMaxRowVersion.get()?.v ?? 0),
    errors,
    conflicts: [] as any[],
  };
}

function upsertRows(table: string, rows: any[]) {
  const keyCol = ensureTableIfMissing(table, DEFAULT_KEY);

  const sample: Record<string, any> = {};
  for (const r of rows)
    for (const [k, v] of Object.entries(r))
      if (k !== keyCol && !(k in sample)) sample[k] = v;
  ensureColumnsForSamples(table, sample, keyCol);

  const cells: CellEditInput[] = [];
  for (const r of rows) {
    const rk = (r as any)[keyCol];
    if (rk === undefined || rk === null) continue;
    for (const [col, val] of Object.entries(r)) {
      if (col === keyCol) continue;
      cells.push({ table, row_key: String(rk), column: col, value: val });
    }
  }
  return upsertCells(cells);
}

function deleteRow(table: string, row_key: string) {
  const keyCol = getKeyCol(table);
  const rv = nextRowVersion();
  const now = Date.now();
  db.prepare(`UPDATE "${table}" SET deleted=1, row_version=?, updated_at=? WHERE "${keyCol}"=?`)
    .run(rv, now, row_key);

  const rowFull = db.prepare(`SELECT * FROM "${table}" WHERE "${keyCol}"=?`).get(row_key) as any;
  const patch = {
    table, row_key: String(row_key), row_version: rowFull.row_version, deleted: true, cells: rowToCells(rowFull, keyCol)
  };
  pubsub.publish("events", { max_row_version: rv, patches: [patch] });
  return true;
}

// ========================= GraphQL schema =========================
const Query = new GraphQLObjectType({
  name: "Query",
  fields: {
    ping: { // 연결상태 확인용(부작용 없음)
      type: new GraphQLNonNull(GraphQLInt),
      resolve: () => Date.now()
    },
    rows: {
      args: { since_version: { type: new GraphQLNonNull(GraphQLInt) } },
      type: new GraphQLNonNull(GqlRowsResult),
      resolve: (_src, { since_version }: { since_version: number }) => {
        const tables = db.prepare(`SELECT table_name FROM _schema`).all() as any[];
        const patches = tables.flatMap(t => readSince(String(t.table_name), since_version));
        const max = Number(getMaxRowVersion.get()?.v ?? 0);
        return { max_row_version: max, patches };
      },
    },
    presence: {
      type: new GraphQLNonNull(new GraphQLList(GqlPresence)),
      resolve: () => db.prepare(`SELECT nickname, sheet, cell, updated_at FROM _presence ORDER BY nickname`).all(),
    },
    audit_log: {
      args: { since_version: { type: GraphQLInt } },
      type: new GraphQLList(GqlAudit),
      resolve: (_src, { since_version }: { since_version?: number }) => {
        if (since_version && since_version > 0) {
          return db.prepare<[number], any>(`SELECT ts,user,table_name AS table,row_key,column_name AS column,old_value,new_value,row_version
                              FROM _audit WHERE row_version > ? ORDER BY id`).all(since_version);
        }
        return db.prepare(`SELECT ts,user,table_name AS table,row_key,column_name AS column,old_value,new_value,row_version
                           FROM _audit ORDER BY id`).all();
      }
    },
    exportDatabase: {
      type: GraphQLString,
      resolve: () => fs.readFileSync(DB_PATH).toString("base64")
    },
    meta: {
      type: GraphQLJSON,
      resolve: () => {
        const m = db.prepare(`SELECT k,v FROM _meta`).all() as any[];
        const s = db.prepare(`SELECT table_name,key_column FROM _schema`).all() as any[];
        return { meta: Object.fromEntries(m.map(r => [r.k, r.v])), schema: s };
      }
    },
    tableColumns: {
      args: { table: { type: new GraphQLNonNull(GraphQLString) } },
      type: new GraphQLNonNull(new GraphQLList(new GraphQLNonNull(GqlColumnInfo))),
      resolve: (_src, { table }: { table: string }) => {
        const rows = db.prepare(`PRAGMA table_info("${table}")`).all() as any[];
        return rows.map(r => ({
          name: String(r.name),
          type: r.type ? String(r.type) : null,
          notnull: Number(r.notnull) === 1,
          pk: Number(r.pk) >= 1
        }));
      }
    }
  }
});

const Mutation = new GraphQLObjectType({
  name: "Mutation",
  fields: {
    upsertCells: {
      args: {
        cells: { type: new GraphQLNonNull(new GraphQLList(new GraphQLNonNull(CellEditInputType))) }
      },
      type: new GraphQLNonNull(GqlUpsertResult),
      resolve: (_src, { cells }: { cells: CellEditInput[] }) => upsertCells(cells)
    },
    upsertRows: {
      args: {
        table: { type: new GraphQLNonNull(GraphQLString) },
        rows: { type: new GraphQLNonNull(new GraphQLList(GraphQLJSON)) }
      },
      type: new GraphQLNonNull(GqlUpsertResult),
      resolve: (_src, { table, rows }: { table: string; rows: any[] }) => upsertRows(table, rows)
    },
    createTable: {
      args: { table: { type: new GraphQLNonNull(GraphQLString) }, key: { type: new GraphQLNonNull(GraphQLString) } },
      type: new GraphQLNonNull(GqlOk),
      resolve: (_src, { table, key }: { table: string; key: string }) => {
        ensureTable(table, key);
        return { ok: true };
      }
    },
    addColumns: {
      args: {
        table: { type: new GraphQLNonNull(GraphQLString) },
        columns: { type: new GraphQLNonNull(new GraphQLList(new GraphQLNonNull(ColumnDefInputType))) }
      },
      type: new GraphQLNonNull(GqlOk),
      resolve: (_src, { table, columns }: { table: string; columns: ColumnDefInput[] }) => {
        addColumns(table, columns);
        return { ok: true };
      }
    },
    presenceTouch: {
      args: {
        nickname: { type: new GraphQLNonNull(GraphQLString) },
        sheet: { type: GraphQLString },
        cell: { type: GraphQLString }
      },
      type: new GraphQLNonNull(GqlOk),
      resolve: (_src, { nickname, sheet, cell }: { nickname: string; sheet?: string; cell?: string }) => {
        const now = Date.now();
        db.prepare(`INSERT INTO _presence(nickname,sheet,cell,updated_at)
                    VALUES(?,?,?,?)
                    ON CONFLICT(nickname) DO UPDATE SET sheet=excluded.sheet, cell=excluded.cell, updated_at=excluded.updated_at`)
          .run(nickname, sheet ?? null, cell ?? null, now);
        return { ok: true };
      }
    },
    acquireLock: {
      args: { cell: { type: new GraphQLNonNull(GraphQLString) }, by: { type: new GraphQLNonNull(GraphQLString) } },
      type: new GraphQLNonNull(GqlOk),
      resolve: (_src, { cell, by }: { cell: string; by: string }) => {
        const now = Date.now();
        const ttlMs = 10000;
        const row = db.prepare<[string], any>(`SELECT by,ts FROM _locks WHERE cell=?`).get(cell);
        if (row) {
          if (now - Number(row.ts) > ttlMs) {
            db.prepare(`UPDATE _locks SET by=?, ts=? WHERE cell=?`).run(by, now, cell);
            return { ok: true };
          }
          return { ok: false };
        } else {
          db.prepare(`INSERT INTO _locks(cell,by,ts) VALUES(?,?,?)`).run(cell, by, now);
          return { ok: true };
        }
      }
    },
    releaseLocksBy: {
      args: { by: { type: new GraphQLNonNull(GraphQLString) } },
      type: new GraphQLNonNull(GqlOk),
      resolve: (_src, { by }: { by: string }) => {
        db.prepare<[string]>(`DELETE FROM _locks WHERE by=?`).run(by);
        return { ok: true };
      }
    },
    dropColumns: {
      args: {
        table: { type: new GraphQLNonNull(GraphQLString) },
        names: { type: new GraphQLNonNull(new GraphQLList(new GraphQLNonNull(GraphQLString))) }
      },
      type: new GraphQLNonNull(GqlOk),
      resolve: (_src, { table, names }: { table: string; names: string[] }) => {
        if (!names.length) return { ok: true };
        const tx = db.transaction(() => {
          for (const n of names) {
            // SQLite 3.35+ 지원: DROP COLUMN
            db.exec(`ALTER TABLE "${table}" DROP COLUMN "${n}"`);
          }
        });
        tx();
        return { ok: true };
      }
    },
  }
});

const Subscription = new GraphQLObjectType({
  name: "Subscription",
  fields: {
    events: {
      type: new GraphQLObjectType({
        name: "ServerEvent",
        fields: {
          max_row_version: { type: new GraphQLNonNull(GraphQLInt) },
          patches: { type: new GraphQLList(GqlPatch) },
        }
      }),
      subscribe: () => pubsub.subscribe("events"),
      resolve: (p: any) => p,
    }
  }
});

const schema = new GraphQLSchema({ query: Query, mutation: Mutation, subscription: Subscription });

// ========================= Server bootstrap =========================
const yoga = createYoga({
  schema,
  graphqlEndpoint: "/graphql",
  maskedErrors: false,
  landingPage: true,
  graphiql: { subscriptionsProtocol: "WS" },
});

const server = createServer(yoga);
const port = Number(process.env.PORT ?? 4000);

server.listen(port, () => {
  console.log(`✅ XQLite GraphQL up`);
  console.log(`   HTTP/UI : http://localhost:${port}/graphql`);
  console.log(`   WS      : ws://localhost:${port}/graphql`);
});
