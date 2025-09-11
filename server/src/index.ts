import express from "express";
import cors from "cors";
import helmet from "helmet";
import bodyParser from "body-parser";
import { ApolloServer } from "@apollo/server";
import { expressMiddleware } from "@as-integrations/express5";
import { GraphQLScalarType, Kind } from "graphql";
import GraphQLJSON from "graphql-type-json";
import Database from "better-sqlite3";

const PORT = parseInt(process.env.PORT || "8000", 10);
const API_KEY = process.env.EXCEL_SQLITE_API_KEY || "devkey";
const DB_PATH = process.env.DB_PATH || "db.sqlite";
const PRESENCE_TTL_SEC = 10;

// ───────── SQLite 연결/초기화 ─────────
const db = new Database(DB_PATH);
db.pragma("journal_mode = WAL");
db.pragma("synchronous = NORMAL");
db.pragma("foreign_keys = ON");
db.pragma("busy_timeout = 5000");

function ident(name: string): string {
    if (!/^[A-Za-z0-9_]+$/.test(name)) {
        const e = new Error(`invalid identifier: ${name}`) as any;
        e.status = 400;
        throw e;
    }
    return name;
}

function ensureAudit() {
    db.exec(`
    CREATE TABLE IF NOT EXISTS audit_log(
      id INTEGER PRIMARY KEY,
      ts TEXT NOT NULL,
      actor TEXT,
      table_name TEXT NOT NULL,
      action TEXT NOT NULL,
      row_id INTEGER,
      detail TEXT
    );
  `);
}
function audit(actor: string, table: string, action: string, rowId: number | null, detail: unknown) {
    const ts = new Date().toISOString();
    db.prepare(
        `INSERT INTO audit_log(ts,actor,table_name,action,row_id,detail)
     VALUES(?,?,?,?,?,?)`
    ).run(ts, actor || "excel", table, action, rowId ?? null, JSON.stringify(detail ?? {}));
}
function ensurePresence() {
    db.exec(`
    CREATE TABLE IF NOT EXISTS presence(
      id INTEGER PRIMARY KEY,
      user TEXT NOT NULL,
      table_name TEXT NOT NULL,
      cell_addr TEXT,
      row_id INTEGER,
      col_name TEXT,
      ts REAL NOT NULL
    );
    CREATE INDEX IF NOT EXISTS ix_presence_recent ON presence(table_name, ts);
  `);
}
ensureAudit();
ensurePresence();

// ───────── GraphQL SDL ─────────
const typeDefs = /* GraphQL */ `
  scalar JSON
  scalar Any

  type ColumnInfo { cid: Int, name: String, type: String, notnull: Int, dflt_value: String, pk: Int }
  type IndexInfo { seq: Int, name: String, unique: Int, origin: String, partial: Int }
  type Meta { columns: [ColumnInfo!]!, indexes: [IndexInfo!]!, max_row_version: Int! }

  input ColumnDefInput {
    name: String!
    type: String!
    not_null: Boolean
    default: String
    pk: Boolean
    unique: Boolean
  }
  input CreateTableInput {
    table: String!
    columns: [ColumnDefInput!]!
    with_meta: Boolean = true
  }
  input AddColumnsInput {
    table: String!
    columns: [ColumnDefInput!]!
  }
  input AddIndexInput {
    table: String!, name: String!, columns: [String!]!, unique: Boolean = false
  }

  type UpsertResult { id: Int, status: String!, db_version: Int, message: String }
  type UpsertPayload { results: [UpsertResult!]!, snapshot: [JSON!]! }

  input RowIn {
    id: Int
    row_version: Int
    data: JSON!
  }

  type Presence {
    user: String!
    table_name: String!
    cell_addr: String
    row_id: Int
    col_name: String
    ts: Float!
  }

  type Query {
    meta(table: String!): Meta!
    rows(
      table: String!,
      since_version: Int,
      whereRaw: String,
      orderBy: String,
      limit: Int,
      offset: Int,
      include_deleted: Boolean = false
    ): [JSON!]!
    presence(table: String!): [Presence!]!
  }

  type Mutation {
    createTable(input: CreateTableInput!): Boolean!
    addColumns(input: AddColumnsInput!): Boolean!
    addIndex(input: AddIndexInput!): Boolean!

    upsertRows(table: String!, actor: String, rows: [RowIn!]!): UpsertPayload!
    deleteRows(table: String!, actor: String, ids: [Int!]!, mode: String = "soft"): Boolean!

    presenceHeartbeat(user: String!, table: String!, cell_addr: String, row_id: Int, col_name: String): Boolean!
  }
`;

// ───────── Scalars ─────────
const AnyScalar = new GraphQLScalarType({
    name: "Any",
    description: "Any scalar",
    serialize: (v) => v,
    parseValue: (v) => v,
    parseLiteral(ast) {
        switch (ast.kind) {
            case Kind.STRING: return ast.value;
            case Kind.BOOLEAN: return ast.value;
            case Kind.INT: return parseInt(ast.value, 10);
            case Kind.FLOAT: return parseFloat(ast.value);
            default: return null;
        }
    }
});

// ───────── Resolvers ─────────
const resolvers = {
    JSON: GraphQLJSON,
    Any: AnyScalar,

    Query: {
        meta: (_: unknown, { table }: { table: string }, ctx: any) => {
            ctx.requireKey();
            const t = ident(table);
            const cols = db.prepare(`PRAGMA table_info(${t});`).all();
            const idxs = db.prepare(`PRAGMA index_list(${t});`).all();
            const rv = db.prepare(`SELECT MAX(row_version) AS maxver FROM ${t};`).get() as any;
            return { columns: cols, indexes: idxs, max_row_version: rv?.maxver || 0 };
        },

        rows: (_: unknown, args: any, ctx: any) => {
            ctx.requireKey();
            const t = ident(args.table);
            const includeDeleted = !!args.include_deleted;
            const since = args.since_version ?? null;

            const where: string[] = [];
            const params: any[] = [];
            if (!includeDeleted) where.push("deleted=0");
            if (since != null) { where.push("row_version > ?"); params.push(since); }

            if (args.whereRaw) {
                const w = String(args.whereRaw);
                if (!/^[\w\s<>=!'.%()\-+/*,]+$/i.test(w)) throw new Error("whereRaw rejected");
                where.push(w);
            }
            const whereSQL = where.length ? ` WHERE ${where.join(" AND ")}` : "";

            let orderSQL = " ORDER BY id";
            if (args.orderBy) {
                const ob = String(args.orderBy);
                if (!/^[\w\s,._-]+$/i.test(ob)) throw new Error("orderBy rejected");
                orderSQL = ` ORDER BY ${ob}`;
            }

            let limitSQL = "";
            if (args.limit != null) limitSQL = ` LIMIT ${parseInt(args.limit, 10)}`;
            let offsetSQL = "";
            if (args.offset != null) offsetSQL = ` OFFSET ${parseInt(args.offset, 10)}`;

            const rows = db.prepare(`SELECT * FROM ${t}${whereSQL}${orderSQL}${limitSQL}${offsetSQL};`).all(...params);
            return rows;
        },

        presence: (_: unknown, { table }: { table: string }, ctx: any) => {
            ctx.requireKey();
            const t = ident(table);
            const now = Date.now() / 1000;
            const rows = db.prepare(`
        SELECT user, table_name, cell_addr, row_id, col_name, ts
        FROM presence
        WHERE table_name=? AND ts > ?
        ORDER BY ts DESC
      `).all(t, now - PRESENCE_TTL_SEC);
            return rows;
        }
    },

    Mutation: {
        createTable: (_: unknown, { input }: any, ctx: any) => {
            ctx.requireKey();
            const { table, columns, with_meta = true } = input;
            const t = ident(table);

            const exists = db.prepare(`SELECT 1 FROM sqlite_master WHERE type='table' AND name=?`).get(t);
            if (exists) throw new Error(`table '${t}' already exists`);

            const userCols = (columns as any[]).map((c) => {
                const n = ident(c.name);
                let sql = `${n} ${c.type || "TEXT"}`;
                if (c.not_null) sql += " NOT NULL";
                if (c.default != null) sql += ` DEFAULT ${c.default}`;
                if (c.unique) sql += " UNIQUE";
                if (c.pk) sql += " PRIMARY KEY";
                return sql;
            });
            const hasPk = !!(columns as any[]).find((c) => c.pk);

            const meta: string[] = [];
            if (with_meta) {
                if (!hasPk) meta.push("id INTEGER PRIMARY KEY");
                meta.push(
                    "row_version INTEGER NOT NULL DEFAULT 0",
                    "updated_at TEXT NOT NULL DEFAULT (strftime('%Y-%m-%dT%H:%M:%fZ','now'))",
                    "deleted INTEGER NOT NULL DEFAULT 0"
                );
            }
            db.exec(`CREATE TABLE ${t} (${[...meta, ...userCols].join(", ")});`);
            audit("ddl", t, "SCHEMA", null, { op: "create_table", with_meta });
            return true;
        },

        addColumns: (_: unknown, { input }: any, ctx: any) => {
            ctx.requireKey();
            const { table, columns } = input;
            const t = ident(table);
            const exists = db.prepare(`SELECT 1 FROM sqlite_master WHERE type='table' AND name=?`).get(t);
            if (!exists) throw new Error(`table '${t}' not found`);

            for (const c of columns as any[]) {
                const n = ident(c.name);
                let sql = `${n} ${c.type || "TEXT"}`;
                if (c.not_null) {
                    if (c.default == null) throw new Error(`NOT NULL ADD requires default for ${n}`);
                    sql += ` NOT NULL DEFAULT ${c.default}`;
                } else if (c.default != null) {
                    sql += ` DEFAULT ${c.default}`;
                }
                db.exec(`ALTER TABLE ${t} ADD COLUMN ${sql}`);
                audit("ddl", t, "SCHEMA", null, { op: "add_column", col: n });
            }
            return true;
        },

        addIndex: (_: unknown, { input }: any, ctx: any) => {
            ctx.requireKey();
            const { table, name, columns, unique = false } = input;
            const t = ident(table), ix = ident(name);
            if (!Array.isArray(columns) || columns.length === 0) throw new Error("columns required");
            const cols = (columns as string[]).map(ident).join(", ");
            const uq = unique ? "UNIQUE " : "";
            db.exec(`CREATE ${uq}INDEX IF NOT EXISTS ${ix} ON ${t} (${cols});`);
            audit("ddl", t, "SCHEMA", null, { op: "add_index", name: ix, columns, unique });
            return true;
        },

        upsertRows: (_: unknown, { table, actor = "excel", rows }: any, ctx: any) => {
            ctx.requireKey();
            const t = ident(table);
            if (!Array.isArray(rows) || rows.length === 0) throw new Error("rows required");

            const colNames = db.prepare(`PRAGMA table_info(${t});`).all().map((r: any) => r.name as string);
            const meta = new Set(["id", "row_version", "updated_at", "deleted"]);
            const userCols = colNames.filter((c) => !meta.has(c));

            const selectVer = db.prepare(`SELECT row_version, deleted FROM ${t} WHERE id=?`);
            const insertAuto = db.prepare(`INSERT INTO ${t} (${userCols.join(",")}) VALUES (${userCols.map(() => "?").join(",")});`);
            const insertWithId = db.prepare(`INSERT INTO ${t} (id, ${userCols.join(",")}) VALUES (?, ${userCols.map(() => "?").join(",")});`);
            const updateStmt = db.prepare(
                `UPDATE ${t} SET ${userCols.map((c) => `${c}=?`).join(", ")},
         row_version=row_version+1, updated_at=strftime('%Y-%m-%dT%H:%M:%fZ','now') WHERE id=?;`
            );
            const selectAll = db.prepare(`SELECT * FROM ${t} ORDER BY id;`);

            const trx = db.transaction((arr: any[]) => {
                const results: any[] = [];
                for (const r of arr) {
                    try {
                        if (r.id == null) {
                            insertAuto.run(...userCols.map((c) => (r.data ?? {})[c]));
                            const newId = (db.prepare("SELECT last_insert_rowid() AS id;").get() as any).id as number;
                            audit(actor, t, "INSERT", newId, { data: r.data });
                            results.push({ id: newId, status: "ok" });
                        } else {
                            const cur = selectVer.get(r.id) as any;
                            if (!cur) {
                                insertWithId.run(r.id, ...userCols.map((c) => (r.data ?? {})[c]));
                                audit(actor, t, "INSERT", r.id, { data: r.data });
                                results.push({ id: r.id, status: "ok" });
                            } else {
                                const dbVer = Number(cur.row_version);
                                if (r.row_version == null || Number(r.row_version) !== dbVer) {
                                    results.push({ id: r.id, status: "conflict", db_version: dbVer });
                                } else {
                                    updateStmt.run(...userCols.map((c) => (r.data ?? {})[c]), r.id);
                                    audit(actor, t, "UPDATE", r.id, { data: r.data });
                                    results.push({ id: r.id, status: "ok" });
                                }
                            }
                        }
                    } catch (err: any) {
                        results.push({ id: r.id ?? null, status: "error", message: String(err?.message || err) });
                    }
                }
                const snapshot = selectAll.all();
                return { results, snapshot };
            });

            return trx(rows);
        },

        deleteRows: (_: unknown, { table, actor = "excel", ids, mode = "soft" }: any, ctx: any) => {
            ctx.requireKey();
            const t = ident(table);
            if (!Array.isArray(ids) || ids.length === 0) return true;

            const trx = db.transaction(() => {
                if (mode === "soft") {
                    const up = db.prepare(
                        `UPDATE ${t} SET deleted=1, row_version=row_version+1,
             updated_at=strftime('%Y-%m-%dT%H:%M:%fZ','now') WHERE id=?;`
                    );
                    for (const id of ids) { up.run(id); audit(actor, t, "DELETE", id, { mode: "soft" }); }
                } else {
                    const del = db.prepare(`DELETE FROM ${t} WHERE id=?;`);
                    for (const id of ids) { del.run(id); audit(actor, t, "DELETE", id, { mode: "hard" }); }
                }
            });
            trx();
            return true;
        },

        presenceHeartbeat: (_: unknown, { user, table, cell_addr, row_id, col_name }: any, ctx: any) => {
            ctx.requireKey();
            const t = ident(table);
            const now = Date.now() / 1000;
            db.prepare(
                `INSERT INTO presence(user, table_name, cell_addr, row_id, col_name, ts)
         VALUES(?,?,?,?,?,?)`
            ).run(String(user), t, cell_addr ?? null, row_id ?? null, col_name ?? null, now);
            db.prepare(`DELETE FROM presence WHERE ts < ?`).run(now - 3600);
            return true;
        }
    }
};

// ───────── 서버 기동 ─────────
async function main() {
    const app = express();
    app.use(helmet());
    app.use(cors());
    app.use(bodyParser.json({ limit: "5mb" }));

    const apollo = new ApolloServer({ typeDefs, resolvers });
    await apollo.start();

    app.use(
        "/graphql",
        expressMiddleware(apollo, {
            context: async ({ req }: { req: express.Request }) => ({
                requireKey: () => {
                    if (req.get("X-API-Key") !== API_KEY) throw new Error("invalid api key");
                }
            })
        })
    );

    app.listen(PORT, () => {
        console.log(`[excel-sqlite-gql-ts] http://localhost:${PORT}/graphql`);
        console.log(`API key: ${API_KEY}`);
    });
}
main();
