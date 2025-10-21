// src/index.ts
//
// XQLite GraphQL Server — Pure Main DB + Admin ATTACH + Row-State Cache
// - main(<project>.sqlite): 순수 데이터만 (id + 비즈니스 컬럼들)
// - admin(<project>.admin.sqlite):
//     _meta(max_row_version), _events(row_version,ts,table,row_key,deleted,cells),
//     _row_state(table,row_key,row_version,ts,deleted),
//     _audit_log, _presence, _locks
// - upsert/delete: 이벤트 기록 + row_state 즉시 갱신 (O(1) 최신상태 조회)
// - rowsSnapshot: main ⨝ row_state 를 서버에서 조인해 스냅샷 제공
// - project: null/"" -> "default"
// - “_” 프리픽스 사용자 테이블 자동 이관(main → admin)
//
// deps:
//   npm i graphql graphql-yoga graphql-type-json better-sqlite3
//   npm i -D ts-node typescript @types/node
//
// run:
//   ts-node --transpile-only src/index.ts
//

import fs from 'fs';
import path from 'path';
import Database from 'better-sqlite3';
import { createServer } from 'http';
import { createSchema, createYoga, createPubSub } from 'graphql-yoga';
import { GraphQLJSON } from 'graphql-type-json';

// ────────────────────────────────────────────────────────────
// ENV
// ────────────────────────────────────────────────────────────
const PORT = process.env.PORT ? Number(process.env.PORT) : 4000;
const DATA_DIR = process.env.XQLITE_DATA_DIR || path.resolve(process.cwd(), 'data');
const PROJECT_FALLBACK = 'default';
const READONLY = process.env.XQLITE_READONLY === '1';

const PRESENCE_TTL_MS = 10_000;
const LOCK_TTL_MS = 10_000;
const CLEAN_INTERVAL_MS = 5_000;
const BROADCAST_BACKSCAN = 1024;

// ────────────────────────────────────────────────────────────
// helpers
// ────────────────────────────────────────────────────────────
function nowMs(): number { return Date.now(); }
function normalizeProject(input?: string | null): string { const v = (input ?? '').trim(); return v || PROJECT_FALLBACK; }
function ensureDir(p: string) { if (!fs.existsSync(p)) fs.mkdirSync(p, { recursive: true }); }
function applyPragmasTo(db: Database.Database, schema?: string) {
  const prefix = schema ? `${schema}.` : '';
  db.pragma(`${prefix}journal_mode = WAL`);
  db.pragma(`${prefix}synchronous = NORMAL`);
  db.pragma(`${prefix}busy_timeout = 5000`);
}
function escapeIdent(name: string): string { return `"${String(name).replace(/"/g, '""')}"`; }
function capInt32(n: number): number { if (!Number.isFinite(n)) return 0; return Math.max(Math.min(n | 0, 2147483647), -2147483648); }
function dataPath(project: string) { ensureDir(DATA_DIR); return path.join(DATA_DIR, `${project}.sqlite`); }
function adminPath(project: string) { ensureDir(DATA_DIR); return path.join(DATA_DIR, `${project}.admin.sqlite`); }
function projectFromCtx(ctx: any): string | undefined {
  try { return ctx?.request?.headers?.get?.('x-project') ?? undefined; } catch { return undefined; }
}
function bundleFrom(ctx: any, projectArg?: string | null): DBBundle {
  const p = normalizeProject(projectArg ?? projectFromCtx(ctx) ?? null);
  return openBundle(p);
}

/** storage class mapping */
function sqlType(t?: string | null) {
  const v = (t ?? '').trim().toLowerCase();
  if (!v) return 'TEXT';
  if (v === 'integer' || v === 'int') return 'INTEGER';
  if (v === 'real' || v === 'float' || v === 'double') return 'REAL';
  if (v === 'bool' || v === 'boolean') return 'INTEGER'; // 0/1
  if (v === 'json') return 'TEXT'; // JSON in TEXT (JSON1)
  if (v === 'text' || v === 'string') return 'TEXT';
  return v.toUpperCase();
}

// ────────────────────────────────────────────────────────────
// types/cache
// ────────────────────────────────────────────────────────────
type DBBundle = {
  project: string;
  db: Database.Database; // main connection (admin ATTACH)
  stmts: {
    // admin
    getMetaVersion: Database.Statement;
    setMetaVersion: Database.Statement;
    selectMaxVersion: Database.Statement;

    insertEvent: Database.Statement;
    selectEventsSince: Database.Statement;

    presenceUpsert: Database.Statement;
    presenceListLive: Database.Statement;
    presenceCleanup: Database.Statement;

    lockAcquire: Database.Statement;
    lockReleaseBy: Database.Statement;
    lockCleanup: Database.Statement;

    auditInsert: Database.Statement;
    auditSince: Database.Statement;

    // row_state
    rowStateUpsert: Database.Statement;
    rowStateGetAll: Database.Statement;
    rowStateGetKeys: (count: number) => Database.Statement;

    // data helpers
    selectPk: (table: string) => string;
    tableInfo: (table: string) => Array<{ name: string; type: string | null; notnull: number; pk: number }>;
    selectRowByPk: (table: string, pkName: string, id: string) => any;
  };
};

const dbCache = new Map<string, DBBundle>();

// admin 전용 표준 테이블(절대 main 금지)
const FORBIDDEN_IN_MAIN = new Set<string>(['_meta', '_events', '_presence', '_locks', '_audit_log', '_row_state']);

// ────────────────────────────────────────────────────────────
// schema hygiene & migration
// ────────────────────────────────────────────────────────────
/** admin attach 후 호출: main에 남은 '_' 테이블들을 admin으로 이관 */
function migrateUnderscoreTablesToAdmin(db: Database.Database) {
  // 후보 수집
  const tables = db.prepare(`SELECT name, sql FROM main.sqlite_master WHERE type='table'`).all() as Array<{ name: string; sql: string | null }>;
  const idxs = db.prepare(`SELECT name, tbl_name, sql FROM main.sqlite_master WHERE type='index' AND sql IS NOT NULL`).all() as Array<{ name: string; tbl_name: string; sql: string }>;
  const trgs = db.prepare(`SELECT name, tbl_name, sql FROM main.sqlite_master WHERE type='trigger' AND sql IS NOT NULL`).all() as Array<{ name: string; tbl_name: string; sql: string }>;

  const targets = tables
    .filter(t => t.name.startsWith('_') && !FORBIDDEN_IN_MAIN.has(t.name));

  if (targets.length === 0) return;

  const tx = db.transaction(() => {
    for (const t of targets) {
      const name = t.name;
      const createSQL = (t.sql ?? '').trim();
      if (!createSQL) {
        // 스키마가 없으면 컬럼 프로빙 후 AS SELECT로 생성
        const cols = (db.prepare(`PRAGMA main.table_info(${escapeIdent(name)})`).all() as any[]).map(c => escapeIdent(c.name)).join(',');
        db.exec(`CREATE TABLE IF NOT EXISTS admin.${escapeIdent(name)} AS SELECT ${cols || '*'} FROM main.${escapeIdent(name)} WHERE 0`);
      } else {
        // CREATE TABLE ... → CREATE TABLE admin."name" ...
        // 첫 테이블 식별자만 admin.<ident>로 치환
        const ddl = createSQL.replace(/^\s*CREATE\s+TABLE\s+(IF\s+NOT\s+EXISTS\s+)?((["`\[]?).+?\3)/i,
          (_m, ifexists) => `CREATE TABLE ${ifexists ?? ''}admin.${escapeIdent(name)}`);
        db.exec(ddl);
      }

      // 컬럼 교집합으로 데이터 복사
      const mainCols = (db.prepare(`PRAGMA main.table_info(${escapeIdent(name)})`).all() as Array<{ name: string }>).map(c => c.name);
      const adminCols = (db.prepare(`PRAGMA admin.table_info(${escapeIdent(name)})`).all() as Array<{ name: string }>).map(c => c.name);
      const common = mainCols.filter(c => adminCols.includes(c));
      if (common.length > 0) {
        const colsList = common.map(escapeIdent).join(',');
        db.exec(`INSERT INTO admin.${escapeIdent(name)}(${colsList}) SELECT ${colsList} FROM main.${escapeIdent(name)}`);
      }

      // 인덱스 재생성(있다면)
      const idxOfTable = idxs.filter(ix => ix.tbl_name === name);
      for (const ix of idxOfTable) {
        const idxDDL = ix.sql
          .replace(/^\s*CREATE\s+INDEX\s+(IF\s+NOT\s+EXISTS\s+)?/i, m => `${m}admin.`)
          .replace(new RegExp(`ON\\s+(${escapeRegex(name)}|${escapeRegex('"' + name + '"')})`, 'i'),
            _m => `ON admin.${escapeIdent(name)}`);
        try { db.exec(idxDDL); } catch { /* ignore conflicting names across schemas */ }
      }

      // 트리거는 스키마 명시/테이블 참조가 복잡할 수 있어 경고만 출력
      const trgOfTable = trgs.filter(tr => tr.tbl_name === name);
      if (trgOfTable.length > 0) {
        console.warn(`[WARN] Triggers on ${name} detected; migration skipped. Please recreate them on admin manually.`);
      }

      // 원본 삭제
      db.exec(`DROP TABLE main.${escapeIdent(name)}`);
      console.log(`[MIGRATE] Moved table ${name} → admin.${name}`);
    }
  });
  tx();
}

function escapeRegex(s: string) {
  return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

/** 메인 스키마에 admin 전용 테이블이 남아 있으면 차단 */
function assertDataSchemaIsPure(db: Database.Database) {
  const rows = db.prepare(`SELECT name FROM main.sqlite_master WHERE type='table'`).all() as Array<{ name: string }>;
  const names = new Set(rows.map(r => r.name));
  const found = [...FORBIDDEN_IN_MAIN].filter(n => names.has(n));
  if (found.length) {
    throw new Error(`Admin tables must NOT exist in main schema: ${found.join(', ')}. Use admin.<table> via ATTACH.`);
  }
}

// ────────────────────────────────────────────────────────────
// open bundle (with migration)
// ────────────────────────────────────────────────────────────
function openBundle(projectRaw?: string | null): DBBundle {
  const project = normalizeProject(projectRaw);
  const cached = dbCache.get(project);
  if (cached) return cached;

  // open main (data)
  const db = new Database(dataPath(project), READONLY ? { readonly: true } : {});
  applyPragmasTo(db);

  // ATTACH admin first (migration needs it)
  const ap = adminPath(project).replace(/"/g, '""');
  db.exec(`ATTACH DATABASE "${ap}" AS admin;`);
  applyPragmasTo(db, 'admin');

  // ── admin schema ensure
  db.exec(`
    CREATE TABLE IF NOT EXISTS admin._meta (key TEXT PRIMARY KEY, value TEXT);
    CREATE TABLE IF NOT EXISTS admin._events (
      row_version INTEGER PRIMARY KEY,
      ts         REAL    NOT NULL,
      table_name TEXT    NOT NULL,
      row_key    TEXT    NOT NULL,
      deleted    INTEGER NOT NULL DEFAULT 0,
      cells      TEXT    NOT NULL
    );
    CREATE TABLE IF NOT EXISTS admin._row_state (
      table_name TEXT    NOT NULL,
      row_key    TEXT    NOT NULL,
      row_version INTEGER NOT NULL,
      ts          REAL    NOT NULL,
      deleted     INTEGER NOT NULL DEFAULT 0,
      PRIMARY KEY (table_name, row_key)
    );
    CREATE TABLE IF NOT EXISTS admin._presence (
      nickname TEXT NOT NULL,
      sheet TEXT,
      cell TEXT,
      updated_at REAL NOT NULL,
      PRIMARY KEY (nickname)
    );
    CREATE TABLE IF NOT EXISTS admin._locks (
      cell TEXT PRIMARY KEY,
      by TEXT NOT NULL,
      updated_at REAL NOT NULL
    );
    CREATE TABLE IF NOT EXISTS admin._audit_log (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      ts REAL NOT NULL,
      user TEXT NOT NULL,
      table_name TEXT NOT NULL,
      row_key TEXT NOT NULL,
      column TEXT,
      old_value TEXT,
      new_value TEXT,
      row_version INTEGER NOT NULL
    );
    CREATE INDEX IF NOT EXISTS admin._events_rowver    ON _events(row_version);
    CREATE INDEX IF NOT EXISTS admin._events_tablekey  ON _events(table_name, row_key);
    CREATE INDEX IF NOT EXISTS admin._row_state_table  ON _row_state(table_name);
    CREATE INDEX IF NOT EXISTS admin._presence_updated ON _presence(updated_at);
    CREATE INDEX IF NOT EXISTS admin._locks_updated    ON _locks(updated_at);
    CREATE INDEX IF NOT EXISTS admin._audit_rowver     ON _audit_log(row_version);
  `);

  // ── migrate '_' tables (user-owned) from main → admin
  if (!READONLY) {
    try { migrateUnderscoreTablesToAdmin(db); }
    catch (e) { console.warn('[WARN] migration failed:', (e as any)?.message ?? e); }
  }

  // ── now ensure main is pure
  assertDataSchemaIsPure(db);

  // _meta init
  const getMetaVersion = db.prepare(`SELECT value FROM admin._meta WHERE key='max_row_version'`);
  const setMetaVersion = db.prepare(`
    INSERT INTO admin._meta(key, value) VALUES('max_row_version', @v)
    ON CONFLICT(key) DO UPDATE SET value=@v
  `);
  if (!getMetaVersion.get()) setMetaVersion.run({ v: '0' });

  // prepared helpers
  const stmts = {
    getMetaVersion,
    setMetaVersion,
    selectMaxVersion: db.prepare(`SELECT CAST(value AS INTEGER) AS v FROM admin._meta WHERE key='max_row_version'`),

    insertEvent: db.prepare(`
      INSERT INTO admin._events(row_version, ts, table_name, row_key, deleted, cells)
      VALUES(@row_version, @ts, @table_name, @row_key, @deleted, @cells)
    `),
    selectEventsSince: db.prepare(`
      SELECT row_version, ts, table_name, row_key, deleted, cells
      FROM admin._events
      WHERE row_version > @since
      ORDER BY row_version ASC
    `),

    presenceUpsert: db.prepare(`
      INSERT INTO admin._presence(nickname, sheet, cell, updated_at)
      VALUES(@nickname, @sheet, @cell, @updated_at)
      ON CONFLICT(nickname) DO UPDATE SET
        sheet=excluded.sheet, cell=excluded.cell, updated_at=excluded.updated_at
    `),
    presenceListLive: db.prepare(`
      SELECT nickname, sheet, cell, updated_at
      FROM admin._presence
      WHERE updated_at >= @since
      ORDER BY updated_at DESC
    `),
    presenceCleanup: db.prepare(`DELETE FROM admin._presence WHERE updated_at < @cutoff`),

    lockAcquire: db.prepare(`
      INSERT INTO admin._locks(cell, by, updated_at)
      VALUES(@cell, @by, @updated_at)
      ON CONFLICT(cell) DO UPDATE SET by=excluded.by, updated_at=excluded.updated_at
    `),
    lockReleaseBy: db.prepare(`DELETE FROM admin._locks WHERE by=@by`),
    lockCleanup: db.prepare(`DELETE FROM admin._locks WHERE updated_at < @cutoff`),

    auditInsert: db.prepare(`
      INSERT INTO admin._audit_log(ts, user, table_name, row_key, column, old_value, new_value, row_version)
      VALUES(@ts, @user, @table_name, @row_key, @column, @old_value, @new_value, @row_version)
    `),
    auditSince: db.prepare(`
      SELECT ts, user, table_name, row_key, column, old_value, new_value, row_version
      FROM admin._audit_log
      WHERE row_version > @since
      ORDER BY row_version ASC
    `),

    // row_state
    rowStateUpsert: db.prepare(`
      INSERT INTO admin._row_state(table_name, row_key, row_version, ts, deleted)
      VALUES(@table_name, @row_key, @row_version, @ts, @deleted)
      ON CONFLICT(table_name, row_key) DO UPDATE SET
        row_version=excluded.row_version,
        ts=excluded.ts,
        deleted=excluded.deleted
    `),
    rowStateGetAll: db.prepare(`
      SELECT row_key, row_version, ts, deleted
      FROM admin._row_state WHERE table_name=@table
      ORDER BY row_key
    `),
    rowStateGetKeys: (count: number) => db.prepare(`
      SELECT row_key, row_version, ts, deleted
      FROM admin._row_state
      WHERE table_name=@table AND row_key IN (${Array(count).fill('?').join(',')})
    `),

    // data helpers
    selectPk: (table: string) => {
      const rows = db.prepare(`PRAGMA table_info(${escapeIdent(table)})`).all() as Array<{ name: string; pk: number }>;
      return rows.find(c => (c.pk | 0) > 0)?.name ?? 'id';
    },
    tableInfo: (table: string) => {
      return db.prepare(`PRAGMA table_info(${escapeIdent(table)})`).all() as Array<{
        name: string; type: string | null; notnull: number; pk: number;
      }>;
    },
    selectRowByPk: (table: string, pkName: string, id: string) => {
      return db.prepare(`SELECT * FROM ${escapeIdent(table)} WHERE ${escapeIdent(pkName)}=@id LIMIT 1`).get({ id });
    },
  };

  const bundle: DBBundle = { project, db, stmts };
  dbCache.set(project, bundle);
  return bundle;
}

function closeAll() {
  for (const [, b] of dbCache) { try { b.db.close(); } catch { } }
  dbCache.clear();
}

function nextRowVersion(bundle: DBBundle): number {
  const cur = bundle.stmts.selectMaxVersion.get() as { v: number } | undefined;
  const next = ((cur?.v ?? 0) | 0) + 1;
  bundle.stmts.setMetaVersion.run({ v: String(next) });
  return next;
}

// ────────────────────────────────────────────────────────────
// 온디맨드 컬럼/테이블 (main = 순수 데이터)
// ────────────────────────────────────────────────────────────
const RESERVED_MAIN = new Set(['id']);

function ensureColumns(bundle: DBBundle, table: string, columns: string[]) {
  if (!columns.length) return;
  const info = bundle.stmts.tableInfo(table);
  const existing = new Set(info.map(c => c.name));
  const toAdd = columns.filter(c => c && !existing.has(c) && !RESERVED_MAIN.has(c));
  for (const name of toAdd) bundle.db.exec(`ALTER TABLE ${escapeIdent(table)} ADD COLUMN ${escapeIdent(name)} TEXT`);
}

function ensureTable(bundle: DBBundle, table: string, key: string) {
  bundle.db.exec(`
    CREATE TABLE IF NOT EXISTS ${escapeIdent(table)} (
      ${escapeIdent(key)} TEXT PRIMARY KEY
    )`);
}

// ────────────────────────────────────────────────────────────
// GraphQL Schema
// ────────────────────────────────────────────────────────────
const typeDefs = /* GraphQL */ `
  scalar JSON

  type Patch {
    table: String!
    row_key: String!
    row_version: Int!
    updated_at: Float!
    deleted: Int!
    cells: JSON!
  }

  type RowsResult {
    max_row_version: Int!
    patches: [Patch!]!
  }

  type UpsertResult {
    max_row_version: Int!
    errors: [String!]
    conflicts: [Conflict!]
  }

  type Conflict {
    table: String!
    row_key: String!
    column: String!
    message: String
  }

  type PresenceItem {
    nickname: String!
    sheet: String
    cell: String
    updated_at: Float
  }

  type ColumnInfo {
    name: String!
    type: String!
    notnull: Boolean!
    pk: Boolean!
  }

  type Ok { ok: Boolean! }

  type Audit {
    ts: Float!
    user: String!
    table: String!
    row_key: String!
    column: String
    old_value: String
    new_value: String
    row_version: Int!
  }

  type RowState {
    row_key: String!
    row_version: Int!
    updated_at: Float!
    deleted: Int!
  }

  input CellEditInput {
    table: String!
    row_key: String!
    column: String!
    value: String
  }

  input ColumnDefInput {
    name: String!
    type: String
    notNull: Boolean
    check: String
  }

  input RenameDefInput {
    from: String!
    to: String!
  }

  input AlterDefInput {
    name: String!
    toType: String
    toNotNull: Boolean
    toCheck: String
  }

  type Query {
    ping: Float!
    rows(since_version: Int, table: String, project: String): RowsResult!
    rowsSnapshot(table: String!, include_deleted: Boolean, project: String): [JSON!]!
    rowState(table: String!, keys: [String!], project: String): [RowState!]!
    meta(project: String): JSON
    audit_log(since_version: Int, project: String): [Audit!]!
    presence(project: String): [PresenceItem!]!
    tableColumns(table: String!, project: String): [ColumnInfo!]!
    exportDatabase(project: String): String
    health: String!
  }

  type Mutation {
    upsertCells(cells: [CellEditInput!]!, project: String): UpsertResult!
    upsertRows(table: String!, rows: [JSON!]!, project: String): UpsertResult!
    deleteRows(table: String!, keys: [String!]!, hard: Boolean, project: String): Ok!
    createTable(table: String!, key: String!, project: String): Ok!
    addColumns(table: String!, columns: [ColumnDefInput!]!, project: String): Ok!
    dropColumns(table: String!, names: [String!]!, project: String): Ok!
    renameColumns(table: String!, renames: [RenameDefInput!]!, project: String): Ok!
    alterColumns(table: String!, alters: [AlterDefInput!]!, project: String): Ok!
    presenceTouch(nickname: String!, sheet: String, cell: String, project: String): Ok!
    acquireLock(cell: String!, by: String!, project: String): Ok!
    releaseLocksBy(by: String!, project: String): Ok!
    rebuildRowState(table: String, project: String): Ok!
  }

  type Subscription {
    events: RowsResult!
  }
`;

// ────────────────────────────────────────────────────────────
// PubSub
// ────────────────────────────────────────────────────────────
type Patch = { table: string; row_key: string; row_version: number; updated_at: number; deleted: number; cells: Record<string, unknown> };
type RowsResult = { max_row_version: number; patches: Patch[] };

const pubsub = createPubSub<{ 'rows-events': [RowsResult] }>();

async function publishEvents(bundle: DBBundle, since: number) {
  const rows = bundle.stmts.selectEventsSince.all({ since }) as Array<{
    row_version: number; ts: number; table_name: string; row_key: string; deleted: number; cells: string;
  }>;
  const patches: Patch[] = rows.map(r => ({
    table: r.table_name,
    row_key: r.row_key,
    row_version: r.row_version,
    updated_at: r.ts,
    deleted: r.deleted | 0,
    cells: JSON.parse(r.cells || '{}'),
  }));
  const mv = bundle.stmts.selectMaxVersion.get() as { v: number };
  const maxRow = (mv?.v ?? 0) | 0;
  await pubsub.publish('rows-events', { max_row_version: maxRow, patches });
}

// ────────────────────────────────────────────────────────────
// Resolvers
// ────────────────────────────────────────────────────────────
const resolvers = {
  JSON: GraphQLJSON,

  Query: {
    ping: () => nowMs(),

    rows: (_: unknown, args: { since_version?: number; table?: string; project?: string }, ctx: any) => {
      const bundle = bundleFrom(ctx, args.project);
      const since = capInt32(Number(args.since_version ?? 0));
      const list = bundle.stmts.selectEventsSince.all({ since }) as Array<{
        row_version: number; ts: number; table_name: string; row_key: string; deleted: number; cells: string;
      }>;
      const patches: Patch[] = list
        .filter(r => (args.table ? r.table_name === args.table : true))
        .map(r => ({
          table: r.table_name, row_key: r.row_key, row_version: r.row_version,
          updated_at: r.ts, deleted: r.deleted | 0, cells: JSON.parse(r.cells || '{}'),
        }));
      const mv = bundle.stmts.selectMaxVersion.get() as { v: number };
      const max = (mv?.v ?? 0) | 0;
      return { max_row_version: max, patches };
    },

    rowsSnapshot: (_: unknown, args: { table: string; include_deleted?: boolean; project?: string }, ctx: any) => {
      const bundle = bundleFrom(ctx, args.project);
      const info = bundle.stmts.tableInfo(args.table);
      if (info.length === 0) return [];
      const pk = info.find(c => (c.pk | 0) > 0)?.name ?? 'id';
      const whereDel = args.include_deleted ? '' : `AND rs.deleted=0`;
      const sql = `
        SELECT m.*, rs.row_version, rs.ts AS updated_at, rs.deleted
        FROM ${escapeIdent(args.table)} m
        LEFT JOIN admin._row_state rs
          ON rs.table_name=@t AND rs.row_key = m.${escapeIdent(pk)}
        WHERE 1=1 ${whereDel}
        ORDER BY m.${escapeIdent(pk)}
      `;
      return bundle.db.prepare(sql).all({ t: args.table }) as any[];
    },

    rowState: (_: unknown, args: { table: string; keys?: string[]; project?: string }, ctx: any) => {
      const bundle = bundleFrom(ctx, args.project);
      if (!args.keys || args.keys.length === 0) {
        const all = bundle.stmts.rowStateGetAll.all({ table: args.table }) as Array<{ row_key: string; row_version: number; ts: number; deleted: number }>;
        return all.map(r => ({ row_key: r.row_key, row_version: r.row_version | 0, updated_at: r.ts, deleted: r.deleted | 0 }));
      } else {
        const stmt = bundle.stmts.rowStateGetKeys(args.keys.length);
        const list = stmt.all(args.table, ...args.keys) as Array<{ row_key: string; row_version: number; ts: number; deleted: number }>;
        return list.map(r => ({ row_key: r.row_key, row_version: r.row_version | 0, updated_at: r.ts, deleted: r.deleted | 0 }));
      }
    },

    meta: (_: unknown, args: { project?: string }, ctx: any) => {
      const bundle = bundleFrom(ctx, args.project);
      const mv = bundle.stmts.selectMaxVersion.get() as { v: number };
      const tables = bundle.db
        .prepare(`
          SELECT name FROM sqlite_master
          WHERE type='table' AND name NOT LIKE 'sqlite_%' AND name NOT LIKE '_%'
          ORDER BY name`)
        .all() as Array<{ name: string }>;
      const schema = tables.map(t => {
        const info = bundle.stmts.tableInfo(t.name);
        const pk = info.find(c => (c.pk | 0) > 0)?.name ?? 'id';
        return { table_name: t.name, key_column: pk };
      });
      const max = (mv?.v ?? 0) | 0;
      return { meta: { max_row_version: String(max) }, schema };
    },

    audit_log: (_: unknown, args: { since_version?: number; project?: string }, ctx: any) => {
      const bundle = bundleFrom(ctx, args.project);
      const since = capInt32(Number(args.since_version ?? 0));
      const rows = bundle.stmts.auditSince.all({ since }) as Array<{
        ts: number; user: string; table_name: string; row_key: string; column: string | null;
        old_value: string | null; new_value: string | null; row_version: number;
      }>;
      return rows.map(r => ({
        ts: r.ts, user: r.user, table: r.table_name, row_key: r.row_key,
        column: r.column, old_value: r.old_value, new_value: r.new_value,
        row_version: r.row_version | 0,
      }));
    },

    presence: (_: unknown, args: { project?: string }, ctx: any) => {
      const bundle = bundleFrom(ctx, args.project);
      const since = nowMs() - PRESENCE_TTL_MS;
      const rows = bundle.stmts.presenceListLive.all({ since }) as Array<{
        nickname: string; sheet: string | null; cell: string | null; updated_at: number;
      }>;
      return rows.map(r => ({ nickname: r.nickname, sheet: r.sheet, cell: r.cell, updated_at: r.updated_at }));
    },

    tableColumns: (_: unknown, args: { table: string; project?: string }, ctx: any) => {
      const bundle = bundleFrom(ctx, args.project);
      const info = bundle.stmts.tableInfo(args.table);
      return info.map(c => ({ name: c.name, type: c.type ?? '', notnull: !!(c.notnull | 0), pk: !!(c.pk | 0) }));
    },

    exportDatabase: (_: unknown, args: { project?: string }, ctx: any) => {
      const bundle = bundleFrom(ctx, args.project);
      const buf = fs.readFileSync(dataPath(bundle.project));
      return buf.toString('base64');
    },

    health: () => 'ok',
  },

  Mutation: {
    presenceTouch: (_: unknown, args: { nickname: string; sheet?: string; cell?: string; project?: string }, ctx: any) => {
      if (READONLY) return { ok: true };
      const bundle = bundleFrom(ctx, args.project);
      bundle.stmts.presenceUpsert.run({
        nickname: args.nickname, sheet: args.sheet ?? null, cell: args.cell ?? null, updated_at: nowMs(),
      });
      return { ok: true };
    },

    acquireLock: (_: unknown, args: { cell: string; by: string; project?: string }, ctx: any) => {
      if (READONLY) return { ok: true };
      const bundle = bundleFrom(ctx, args.project);
      bundle.stmts.lockAcquire.run({ cell: args.cell, by: args.by, updated_at: nowMs() });
      return { ok: true };
    },

    releaseLocksBy: (_: unknown, args: { by: string; project?: string }, ctx: any) => {
      if (READONLY) return { ok: true };
      const bundle = bundleFrom(ctx, args.project);
      bundle.stmts.lockReleaseBy.run({ by: args.by });
      return { ok: true };
    },

    createTable: (_: unknown, args: { table: string; key: string; project?: string }, ctx: any) => {
      if (READONLY) return { ok: true };
      const bundle = bundleFrom(ctx, args.project);
      bundle.db.exec(`
        CREATE TABLE IF NOT EXISTS ${escapeIdent(args.table)} (
          ${escapeIdent(args.key)} TEXT PRIMARY KEY
        )`);
      return { ok: true };
    },

    addColumns: (_: unknown, args: { table: string; columns: Array<{ name: string; type?: string; notNull?: boolean; check?: string }>; project?: string }, ctx: any) => {
      if (READONLY) return { ok: true };
      const bundle = bundleFrom(ctx, args.project);
      const cols = args.columns.map(c => ({
        name: c.name,
        type: sqlType(c.type),
        check: (c.check ?? '').trim(),
      }));
      const exists = bundle.stmts.tableInfo(args.table).map(x => x.name);
      const existSet = new Set(exists);

      const tx = bundle.db.transaction(() => {
        for (const c of cols) {
          if (!c.name || RESERVED_MAIN.has(c.name) || existSet.has(c.name)) continue;
          let ddl = `ALTER TABLE ${escapeIdent(args.table)} ADD COLUMN ${escapeIdent(c.name)} ${c.type}`;
          if (c.check) ddl += ` CHECK(${c.check})`;
          bundle.db.exec(ddl);
          existSet.add(c.name);
        }
      });
      tx();
      return { ok: true };
    },

    dropColumns: (_: unknown, args: { table: string; names: string[]; project?: string }, ctx: any) => {
      if (READONLY) return { ok: true };
      const bundle = bundleFrom(ctx, args.project);
      const info = bundle.stmts.tableInfo(args.table);
      if (!info.length) return { ok: true };
      const pk = info.find(c => (c.pk | 0) > 0)?.name ?? 'id';
      const targets = (args.names ?? []).filter(n => n && n !== pk && !RESERVED_MAIN.has(n));
      if (!targets.length) return { ok: true };

      // 최신 SQLite에서 직접 DROP COLUMN 시도
      let allOk = true;
      for (const n of targets) {
        try {
          bundle.db.exec(`ALTER TABLE ${escapeIdent(args.table)} DROP COLUMN ${escapeIdent(n)}`);
        } catch {
          allOk = false;
          break;
        }
      }
      if (allOk) return { ok: true };

      // 폴백: 재구성
      const cols = bundle.stmts.tableInfo(args.table);
      const kept = cols.filter(c => !targets.includes(c.name));
      const def = kept.map(c => {
        const typ = c.type ? ` ${c.type}` : '';
        const nn = (c.notnull | 0) ? ' NOT NULL' : '';
        const pkd = (c.pk | 0) ? ' PRIMARY KEY' : '';
        return `${escapeIdent(c.name)}${typ}${nn}${pkd}`;
      }).join(', ');
      const tmp = `_tmp_${Date.now().toString(36)}`;
      const colList = kept.map(c => escapeIdent(c.name)).join(', ');

      const tx = bundle.db.transaction(() => {
        bundle.db.exec(`CREATE TABLE ${escapeIdent(tmp)} (${def})`);
        bundle.db.exec(`INSERT INTO ${escapeIdent(tmp)} (${colList}) SELECT ${colList} FROM ${escapeIdent(args.table)}`);
        bundle.db.exec(`DROP TABLE ${escapeIdent(args.table)}`);
        bundle.db.exec(`ALTER TABLE ${escapeIdent(tmp)} RENAME TO ${escapeIdent(args.table)}`);
      });
      tx();
      return { ok: true };
    },

    renameColumns: (_: unknown, args: { table: string; renames: Array<{ from: string; to: string }>; project?: string }, ctx: any) => {
      if (READONLY) return { ok: true };
      const bundle = bundleFrom(ctx, args.project);
      const info = bundle.stmts.tableInfo(args.table);
      if (info.length === 0) return { ok: true };

      const reserved = new Set(['row_version', 'updated_at', 'deleted']); // defensive
      const pk = info.find(c => (c.pk | 0) > 0)?.name;

      const tx = bundle.db.transaction(() => {
        for (const r of args.renames) {
          if (!r?.from || !r?.to) continue;
          if (pk && r.from === pk) continue;
          if (reserved.has(r.from)) continue;
          if (r.from === r.to) continue;
          const sql = `ALTER TABLE ${escapeIdent(args.table)} RENAME COLUMN ${escapeIdent(r.from)} TO ${escapeIdent(r.to)}`;
          bundle.db.exec(sql);
        }
      });
      tx();

      return { ok: true };
    },

    alterColumns: (_: unknown, args: { table: string; alters: Array<{ name: string; toType?: string | null; toNotNull?: boolean | null; toCheck?: string | null }>; project?: string }, ctx: any) => {
      if (READONLY) return { ok: true };
      const bundle = bundleFrom(ctx, args.project);
      const info = bundle.stmts.tableInfo(args.table);
      if (info.length === 0) return { ok: true };

      const wants = new Map<string, { toType?: string | null; toNotNull?: boolean | null; toCheck?: string | null }>();
      for (const a of args.alters ?? []) {
        if (!a?.name) continue;
        wants.set(a.name, { toType: (a.toType ?? null), toNotNull: a.toNotNull ?? null, toCheck: a.toCheck ?? null });
      }
      if (wants.size === 0) return { ok: true };

      rebuildTableWithAlters(bundle, args.table, info, wants);
      return { ok: true };
    },

    deleteRows: (_: unknown, args: { table: string; keys: string[]; hard?: boolean; project?: string }, ctx: any) => {
      if (READONLY) return { ok: true };
      const bundle = bundleFrom(ctx, args.project);
      const info = bundle.stmts.tableInfo(args.table);
      if (info.length === 0) return { ok: true };
      const pk = info.find(c => (c.pk | 0) > 0)?.name ?? 'id';

      const tx = bundle.db.transaction(() => {
        for (const id of args.keys) {
          const rv = nextRowVersion(bundle);
          const ts = nowMs();
          bundle.db.prepare(`DELETE FROM ${escapeIdent(args.table)} WHERE ${escapeIdent(pk)}=@id`).run({ id });
          bundle.stmts.insertEvent.run({
            row_version: rv, ts, table_name: args.table, row_key: id, deleted: 1, cells: JSON.stringify({}),
          });
          bundle.stmts.rowStateUpsert.run({
            table_name: args.table, row_key: id, row_version: rv, ts, deleted: 1,
          });
          bundle.stmts.auditInsert.run({
            ts, user: 'excel', table_name: args.table, row_key: id,
            column: 'deleted', old_value: '0', new_value: args.hard ? 'HARD_DELETE' : '1', row_version: rv,
          });
        }
      });
      tx();

      const mv = bundle.stmts.selectMaxVersion.get() as { v: number };
      const maxRow = (mv?.v ?? 0) | 0;
      publishEvents(bundle, Math.max(0, maxRow - BROADCAST_BACKSCAN)).catch(() => { });
      return { ok: true };
    },

    upsertCells: (_: unknown, args: { cells: Array<{ table: string; row_key: string; column: string; value?: string | null }>; project?: string }, ctx: any) => {
      const bundle = bundleFrom(ctx, args.project);
      if (READONLY) {
        const mv = bundle.stmts.selectMaxVersion.get() as { v: number };
        return { max_row_version: (mv?.v ?? 0) | 0, errors: ['READONLY'], conflicts: null };
      }

      const perTable = new Map<string, Set<string>>();
      for (const e of args.cells) {
        ensureTable(bundle, e.table, 'id');
        const key = e.table;
        if (!perTable.has(key)) perTable.set(key, new Set());
        if (e.column) perTable.get(key)!.add(e.column);
      }
      for (const [table, set] of perTable) ensureColumns(bundle, table, Array.from(set));

      const errors: string[] = [];
      const grouped = new Map<string, Array<{ column: string; value: any }>>();
      const keyMap = new Map<string, { table: string; row_key: string }>();
      for (const c of args.cells) {
        if (!c.table || !c.row_key || !c.column) continue;
        const key = `${c.table}::${c.row_key}`;
        if (!grouped.has(key)) grouped.set(key, []);
        grouped.get(key)!.push({ column: c.column, value: c.value ?? null });
        keyMap.set(key, { table: c.table, row_key: c.row_key });
      }

      const tx = bundle.db.transaction(() => {
        for (const [k, edits] of grouped) {
          const { table, row_key } = keyMap.get(k)!;
          const info = bundle.stmts.tableInfo(table);
          const colSet = new Set(info.map(x => x.name));
          const pk = info.find(c => (c.pk | 0) > 0)?.name ?? 'id';

          const before = bundle.stmts.selectRowByPk(table, pk, row_key) ?? null;

          const cols = Array.from(new Set(edits.map(e => e.column))).filter(c => colSet.has(c));
          const kv: Record<string, any> = {};
          for (const e of edits) if (colSet.has(e.column)) kv[e.column] = e.value;

          const exists = !!before;
          const rv = nextRowVersion(bundle);
          const ts = nowMs();

          if (exists) {
            if (cols.length > 0) {
              const sets = cols.map(c => `${escapeIdent(c)}=@${c}`);
              const sql = `UPDATE ${escapeIdent(table)} SET ${sets.join(', ')} WHERE ${escapeIdent(pk)}=@id`;
              bundle.db.prepare(sql).run({ ...kv, id: row_key });
            }
          } else {
            const allCols = [pk, ...cols];
            const placeholders = allCols.map(c => `@${c}`);
            const p: any = { [pk]: row_key };
            for (const c of cols) p[c] = kv[c];
            const sql = `INSERT INTO ${escapeIdent(table)}(${allCols.map(escapeIdent).join(', ')}) VALUES(${placeholders.join(', ')})`;
            bundle.db.prepare(sql).run(p);
          }

          for (const c of cols) {
            const oldVal = before ? (before[c] != null ? String(before[c]) : null) : null;
            const newVal = kv[c] != null ? String(kv[c]) : null;
            bundle.stmts.auditInsert.run({
              ts, user: 'excel', table_name: table, row_key, column: c, old_value: oldVal, new_value: newVal, row_version: rv,
            });
          }
          bundle.stmts.insertEvent.run({
            row_version: rv, ts, table_name: table, row_key, deleted: 0, cells: JSON.stringify(kv),
          });
          bundle.stmts.rowStateUpsert.run({
            table_name: table, row_key, row_version: rv, ts, deleted: 0,
          });
        }
      });

      try { tx(); } catch (e: any) { errors.push(String(e?.message ?? e)); }

      const mv = bundle.stmts.selectMaxVersion.get() as { v: number };
      const maxRow = (mv?.v ?? 0) | 0;
      publishEvents(bundle, Math.max(0, maxRow - BROADCAST_BACKSCAN)).catch(() => { });
      return { max_row_version: maxRow, errors: errors.length ? errors : null, conflicts: null };
    },

    upsertRows: (_: unknown, args: { table: string; rows: Array<Record<string, any>>; project?: string }, ctx: any) => {
      const bundle = bundleFrom(ctx, args.project);
      if (READONLY) {
        const mv = bundle.stmts.selectMaxVersion.get() as { v: number };
        return { max_row_version: (mv?.v ?? 0) | 0, errors: ['READONLY'], conflicts: null };
      }

      ensureTable(bundle, args.table, 'id');
      const keys = new Set<string>();
      for (const r of args.rows) for (const k of Object.keys(r)) if (k) keys.add(k);
      const info0 = bundle.stmts.tableInfo(args.table);
      const pk = info0.find(c => (c.pk | 0) > 0)?.name ?? 'id';
      keys.delete(pk);
      ensureColumns(bundle, args.table, Array.from(keys));

      const errors: string[] = [];

      const tx2 = bundle.db.transaction(() => {
        const info2 = bundle.stmts.tableInfo(args.table);
        const colSet = new Set(info2.map(x => x.name));

        for (const row of args.rows) {
          const row_key = String(row[pk] ?? '');
          if (!row_key) continue;

          const before = bundle.stmts.selectRowByPk(args.table, pk, row_key) ?? null;
          const cols = Object.keys(row).filter(c => c !== pk && colSet.has(c));
          const kv: Record<string, any> = {};
          for (const c of cols) kv[c] = row[c];

          const exists = !!before;
          const rv = nextRowVersion(bundle);
          const ts = nowMs();

          if (exists) {
            if (cols.length > 0) {
              const sets = cols.map(c => `${escapeIdent(c)}=@${c}`);
              const sql = `UPDATE ${escapeIdent(args.table)} SET ${sets.join(', ')} WHERE ${escapeIdent(pk)}=@id`;
              bundle.db.prepare(sql).run({ ...kv, id: row_key });
            }
          } else {
            const allCols = [pk, ...cols];
            const placeholders = allCols.map(c => `@${c}`);
            const p: any = { [pk]: row_key };
            for (const c of cols) p[c] = kv[c];
            const sql = `INSERT INTO ${escapeIdent(args.table)}(${allCols.map(escapeIdent).join(', ')}) VALUES(${placeholders.join(', ')})`;
            bundle.db.prepare(sql).run(p);
          }

          for (const c of cols) {
            const oldVal = before ? (before[c] != null ? String(before[c]) : null) : null;
            const newVal = kv[c] != null ? String(kv[c]) : null;
            bundle.stmts.auditInsert.run({
              ts, user: 'excel', table_name: args.table, row_key, column: c, old_value: oldVal, new_value: newVal, row_version: rv,
            });
          }

          bundle.stmts.insertEvent.run({
            row_version: rv, ts, table_name: args.table, row_key, deleted: 0, cells: JSON.stringify(kv),
          });
          bundle.stmts.rowStateUpsert.run({
            table_name: args.table, row_key, row_version: rv, ts, deleted: 0,
          });
        }
      });

      try { tx2(); } catch (e: any) { errors.push(String(e?.message ?? e)); }

      const mv = bundle.stmts.selectMaxVersion.get() as { v: number };
      const maxRow = (mv?.v ?? 0) | 0;
      publishEvents(bundle, Math.max(0, maxRow - BROADCAST_BACKSCAN)).catch(() => { });
      return { max_row_version: maxRow, errors: errors.length ? errors : null, conflicts: null };
    },

    rebuildRowState: (_: unknown, args: { table?: string; project?: string }, ctx: any) => {
      if (READONLY) return { ok: true };
      const bundle = bundleFrom(ctx, args.project);

      const tx = bundle.db.transaction(() => {
        if (args.table) {
          bundle.db.prepare(`DELETE FROM admin._row_state WHERE table_name=@t`).run({ t: args.table });
          const cur = bundle.db.prepare(`
            SELECT table_name, row_key,
                   MAX(row_version) AS row_version,
                   MAX(ts) AS ts,
                   (SELECT deleted FROM admin._events e2
                      WHERE e2.table_name=e1.table_name AND e2.row_key=e1.row_key
                      ORDER BY row_version DESC LIMIT 1) AS deleted
            FROM admin._events e1
            WHERE table_name=@t
            GROUP BY table_name, row_key
          `).all({ t: args.table }) as Array<{ table_name: string; row_key: string; row_version: number; ts: number; deleted: number }>;
          const ins = bundle.stmts.rowStateUpsert;
          for (const r of cur) ins.run({ table_name: r.table_name, row_key: r.row_key, row_version: r.row_version, ts: r.ts, deleted: r.deleted | 0 });
        } else {
          bundle.db.exec(`DELETE FROM admin._row_state`);
          const cur = bundle.db.prepare(`
            SELECT table_name, row_key,
                   MAX(row_version) AS row_version,
                   MAX(ts) AS ts,
                   (SELECT deleted FROM admin._events e2
                      WHERE e2.table_name=e1.table_name AND e2.row_key=e1.row_key
                      ORDER BY row_version DESC LIMIT 1) AS deleted
            FROM admin._events e1
            GROUP BY table_name, row_key
          `).all() as Array<{ table_name: string; row_key: string; row_version: number; ts: number; deleted: number }>;
          const ins = bundle.stmts.rowStateUpsert;
          for (const r of cur) ins.run({ table_name: r.table_name, row_key: r.row_key, row_version: r.row_version, ts: r.ts, deleted: r.deleted | 0 });
        }
      });
      tx();
      return { ok: true };
    },
  },

  Subscription: {
    events: {
      subscribe: () => pubsub.subscribe('rows-events'),
      resolve: (payload: RowsResult) => payload,
    },
  },
};

// ────────────────────────────────────────────────────────────
// Server
// ────────────────────────────────────────────────────────────
const schema = createSchema({ typeDefs, resolvers });
const yoga = createYoga({ schema, logging: true, maskedErrors: false });
const httpServer = createServer(yoga);

// cleaners
const cleaners = new Set<NodeJS.Timeout>();
function startCleaners() {
  const t = setInterval(() => {
    for (const [, b] of dbCache) {
      try {
        const now = nowMs();
        b.stmts.presenceCleanup.run({ cutoff: now - PRESENCE_TTL_MS });
        b.stmts.lockCleanup.run({ cutoff: now - LOCK_TTL_MS });
      } catch { /* ignore */ }
    }
  }, CLEAN_INTERVAL_MS);
  cleaners.add(t);
}
function stopCleaners() { for (const t of cleaners) clearInterval(t); cleaners.clear(); }

httpServer.listen(PORT, () => {
  startCleaners();
  console.log(`${new Date().toISOString()} XQLite server ready on :${PORT} (data dir: ${DATA_DIR}) RO=${READONLY ? '1' : '0'}`);
});

// graceful shutdown
function shutdown(code = 0) {
  try { stopCleaners(); } catch { }
  try { httpServer.close(() => { }); } catch { }
  try { closeAll(); } catch { }
  setTimeout(() => process.exit(code), 200).unref();
}
process.on('SIGINT', () => { console.log('SIGINT'); shutdown(0); });
process.on('SIGTERM', () => { console.log('SIGTERM'); shutdown(0); });
process.on('uncaughtException', (err) => { console.error('[uncaughtException]', err?.stack || err); });
process.on('unhandledRejection', (reason) => { console.error('[unhandledRejection]', reason); });
function rebuildTableWithAlters(bundle: DBBundle, table: string, info: { name: string; type: string | null; notnull: number; pk: number; }[], wants: Map<string, { toType?: string | null; toNotNull?: boolean | null; toCheck?: string | null; }>) {
  throw new Error('Function not implemented.');
}

