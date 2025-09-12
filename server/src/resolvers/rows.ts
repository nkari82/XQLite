import { db } from "../db";
import { nextRowVersion } from "../db";

const META_COLS = `id INTEGER PRIMARY KEY, row_version INTEGER NOT NULL, updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP, deleted INTEGER NOT NULL DEFAULT 0`;

function ensureTable(table: string, columns: Record<string, string>) {
    // columns: { "name":"TEXT", "atk":"INTEGER", ... }
    const cols = Object.entries(columns).map(([k, v]) => `${k} ${v}`).join(", ");
    db.exec(`CREATE TABLE IF NOT EXISTS ${table} (${META_COLS}${cols ? ", " + cols : ""})`);
    // 인덱스 예시
    db.exec(`CREATE INDEX IF NOT EXISTS ix_${table}_row_version ON ${table}(row_version)`);
}

export const queryRows = (_: any, args: any) => {
    const { table, since_version, whereRaw, orderBy, limit = 5000, offset = 0, include_deleted = false } = args;
    const conds = [];
    if (!include_deleted) conds.push("deleted=0");
    if (since_version != null) conds.push(`row_version > ${Number(since_version)}`);
    if (whereRaw) conds.push(`(${whereRaw})`);
    const where = conds.length ? `WHERE ${conds.join(" AND ")}` : "";
    const order = orderBy ? `ORDER BY ${orderBy}` : "";
    const rows = db.prepare(`SELECT * FROM ${table} ${where} ${order} LIMIT ? OFFSET ?`).all(limit, offset);
    const maxv = Number((db.prepare(`SELECT value FROM meta WHERE key='max_row_version'`).get() as { value: number | string }).value);
    return { rows, max_row_version: maxv, affected: rows.length, conflicts: [], errors: [] };
};

export const upsertRows = (_: any, { table, rows, actor }: any) => {
    if (!rows?.length) return { rows: [], max_row_version: 0, affected: 0, conflicts: [], errors: [] };
    // 테이블 존재 가정(없다면 사전에 createTable/addColumns로 생성)
    const cols = Object.keys(rows[0]).filter(k => !["id", "row_version", "updated_at", "deleted"].includes(k));
    const placeholders = cols.map(_ => "?").join(",");
    const setClause = cols.map(c => `${c}=?`).join(",");
    const now = new Date().toISOString();

    const insert = db.prepare(`
    INSERT INTO ${table} (${["id", ...cols, "row_version", "updated_at", "deleted"].join(",")})
    VALUES (? , ${placeholders} , ? , ? , 0)
    ON CONFLICT(id) DO UPDATE SET ${setClause}, row_version=excluded.row_version, updated_at=excluded.updated_at
  `);
    const tx = db.transaction(() => {
        let affected = 0;
        for (const r of rows) {
            const rv = nextRowVersion();
            const valsIns = [r.id, ...cols.map(c => r[c]), rv, now];
            insert.run(valsIns);
            affected++;
        }
        db.prepare(`INSERT INTO audit_log(actor,action,table_name,detail) VALUES (?,?,?,?)`)
            .run(actor, "upsert", table, JSON.stringify({ count: rows.length }));
        const maxv = Number((db.prepare(`SELECT value FROM meta WHERE key='max_row_version'`).get() as { value: number | string }).value);
        return { rows: [], max_row_version: maxv, affected, conflicts: [], errors: [] };
    });
    return tx();
};

export const deleteRows = (_: any, { table, ids, actor }: any) => {
    if (!ids?.length) return { rows: [], max_row_version: 0, affected: 0, conflicts: [], errors: [] };
    const now = new Date().toISOString();
    const tx = db.transaction(() => {
        let affected = 0;
        for (const id of ids) {
            const rv = nextRowVersion();
            db.prepare(`UPDATE ${table} SET deleted=1, row_version=?, updated_at=? WHERE id=?`).run(rv, now, id);
            affected++;
        }
        db.prepare(`INSERT INTO audit_log(actor,action,table_name,detail) VALUES (?,?,?,?)`)
            .run(actor, "delete", table, JSON.stringify({ count: ids.length }));
        const maxv = Number((db.prepare(`SELECT value FROM meta WHERE key='max_row_version'`).get() as { value: number | string }).value);
        return { rows: [], max_row_version: maxv, affected, conflicts: [], errors: [] };
    });
    return tx();
};

export const recoverFromExcel = (_: any, { table, rows, schema_hash, actor }: any) => {
    // 1) (선택) schema_hash 비교 후 불일치면 스키마 재생성/보정
    // 2) 테이블 truncate 후 배치 업서트
    const now = new Date().toISOString();
    const tx = db.transaction(() => {
        db.exec(`DELETE FROM ${table}`);
        let affected = 0;
        for (const r of rows) {
            const rv = nextRowVersion();
            const keys = Object.keys(r);
            const cols = keys.join(",");
            const ph = keys.map(_ => "?").join(",");
            const vals = keys.map(k => r[k]);
            db.prepare(`INSERT INTO ${table} (${cols}, row_version, updated_at, deleted) VALUES (${ph}, ?, ?, 0)`)
                .run(...vals, rv, now);
            affected++;
        }
        db.prepare(`INSERT INTO audit_log(actor,action,table_name,detail) VALUES (?,?,?,?)`)
            .run(actor, "recover", table, JSON.stringify({ count: rows.length, schema_hash }));
        const maxv = Number((db.prepare(`SELECT value FROM meta WHERE key='max_row_version'`).get() as { value: number | string }).value);
        return maxv;
    });
    const maxv = tx();
    return true;
};
