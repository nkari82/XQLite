import { db, nextRowVersion } from "../db.js";
import { getTableDef, zodShapeFor, TableDef } from "./registry.js";
import { ERR } from "../errors.js";
import { sanitizeOrderBy } from "../util/sql.js";
import { notifyChange } from "../notifier.js"

type RowData = Record<string, any> & { row_version: number };

const META_FORBID = new Set(["row_version", "updated_at", "deleted"]);
const SQL_FORBID = /\b(ATTACH|PRAGMA|UNION|SELECT|INSERT|UPDATE|DELETE|DROP|ALTER|CREATE|;|--|\/\*)\b/i;

function assertKnownTable(table: string): TableDef {
    const def = getTableDef(table);
    if (!def) throw ERR.VALID(`unknown table ${table}`);
    return def;
}

function validateWhereRaw(def: TableDef, whereRaw?: string) {
    if (!whereRaw) return "";
    if (SQL_FORBID.test(whereRaw)) throw ERR.VALID("whereRaw forbidden keyword");

    // 토큰 분해 후, 컬럼 식별자만 화이트리스트 확인
    const words = whereRaw.match(/[A-Za-z_][A-Za-z0-9_\.]*/g) ?? [];
    const allowed = new Set(def.columns.map(c => c.name).concat(["id", "deleted"]));
    for (const w of words) {
        const lw = w.toLowerCase();
        // SQL 키워드 대충 패스, 컬럼 후보만 검사
        if (["and", "or", "not", "is", "null", "like", "in", "between", "case", "when", "then", "else", "end"].includes(lw)) continue;
        if (/^\d+$/.test(lw)) continue;
        if (!allowed.has(w) && !allowed.has(lw)) {
            throw ERR.VALID(`whereRaw uses unknown identifier: ${w}`);
        }
    }
    return `(${whereRaw})`;
}

export const queryRows = (_: any, args: any) => {
    const { table, since_version, whereRaw, orderBy, limit = 5000, offset = 0, include_deleted = false } = args;
    const def = assertKnownTable(table);

    const conds: string[] = [];
    if (!include_deleted) conds.push("deleted=0");
    if (since_version != null) conds.push(`row_version > ${Number(since_version)}`);
    if (whereRaw) conds.push(validateWhereRaw(def, whereRaw));
    const where = conds.length ? `WHERE ${conds.join(" AND ")}` : "";
    const order = sanitizeOrderBy(orderBy, def);
    const rows = db.prepare(`SELECT * FROM ${table} ${where} ${order} LIMIT ? OFFSET ?`).all(limit, offset);
    const maxv = Number((db.prepare(`SELECT value FROM meta WHERE key='max_row_version'`).get() as { value: string }).value);
    return { rows, max_row_version: maxv, affected: rows.length, conflicts: [], errors: [] };
};

// ───────────────────────────────────────────
// V2: 낙관적 잠금 + 셀 충돌
// ───────────────────────────────────────────
const toNFC = (v: any) => typeof v === "string" ? v.normalize("NFC") : v;

export const upsertRows = (_: any, { table, rows, actor }: any) => {
    const def = assertKnownTable(table);
    const validator = zodShapeFor(def);
    const now = new Date().toISOString();

    const tx = db.transaction(() => {
        let affected = 0;
        const conflicts: any[] = [];

        for (const item of rows ?? []) {
            const { id, base_row_version = 0, data } = item;
            if (typeof id !== "number") throw ERR.VALID("id must be number");

            // 1) 검증 & 정규화
            const norm: any = {};
            for (const [k, v] of Object.entries(data ?? {})) {
                if (META_FORBID.has(k)) continue;
                norm[k] = toNFC(v);
            }
            // 부분유효성(없는 컬럼은 무시)
            const parsed = validator.partial().parse({ id, ...norm });

            // 2) 현재 행과 충돌검사
            const cur = db.prepare(`SELECT * FROM ${table} WHERE id=?`).get(id) as RowData;
            if (cur && base_row_version && cur.row_version > base_row_version) {
                const changed: string[] = [];
                for (const [k, vNew] of Object.entries(parsed)) {
                    if (k === "id") continue;
                    if (cur[k] !== vNew) changed.push(k);
                }
                if (changed.length) {
                    conflicts.push({
                        id,
                        columns: changed,
                        server: changed.reduce((m, c) => (m[c] = cur[c], m), {} as any),
                        server_row_version: cur.row_version
                    });
                    continue; // 충돌 시 스킵
                }
            }

            // 3) 업서트
            const rv = nextRowVersion();
            const keys = Object.keys(parsed).filter(k => k !== "id");
            const cols = ["id", ...keys, "row_version", "updated_at", "deleted"];
            const ph = cols.map(_ => "?").join(",");
            const vals = [id, ...keys.map(k => (parsed as any)[k]), rv, now, 0];

            db.prepare(`
        INSERT INTO ${table} (${cols.join(",")}) VALUES (${ph})
        ON CONFLICT(id) DO UPDATE SET
          ${keys.map(k => `${k}=excluded.${k}`).join(",")},
          row_version=excluded.row_version,
          updated_at=excluded.updated_at
      `).run(...vals);

            // 4) 셀 단위 감사 로그
            if (cur) {
                const diffs = keys.filter(k => cur[k] !== (parsed as any)[k]);
                for (const k of diffs) {
                    db.prepare(`INSERT INTO audit_log(actor, action, table_name, detail) VALUES (?, 'cell_update', ?, ?)`)
                        .run(actor, table, JSON.stringify({ id, column: k, from: cur[k], to: (parsed as any)[k] }));
                }
            }
            affected++;
        }

        const maxv = Number((db.prepare(`SELECT value FROM meta WHERE key='max_row_version'`).get() as { value: string }).value);
        return { rows: [], max_row_version: maxv, affected, conflicts, errors: [] };
    });

    const out = tx();
    notifyChange(table, out.max_row_version);   // ← 커밋 후 알림
    return out;
};

export const deleteRows = (_: any, { table, ids, actor }: any) => {
    assertKnownTable(table);
    const now = new Date().toISOString();
    const tx = db.transaction(() => {
        let affected = 0;
        for (const id of ids ?? []) {
            const rv = nextRowVersion();
            db.prepare(`UPDATE ${table} SET deleted=1, row_version=?, updated_at=? WHERE id=?`).run(rv, now, id);
            affected++;
        }
        const maxv = Number((db.prepare(`SELECT value FROM meta WHERE key='max_row_version'`).get() as { value: string }).value);
        return { rows: [], max_row_version: maxv, affected, conflicts: [], errors: [] };
    });
    const out = tx();
    notifyChange(table, out.max_row_version);
    return out;
};

export const recoverFromExcel = (_: any, { table, rows, schema_hash, actor }: any) => {
    assertKnownTable(table);
    const now = new Date().toISOString();
    const tx = db.transaction(() => {
        db.exec(`DELETE FROM ${table}`);
        for (const r of rows ?? []) {
            const rv = nextRowVersion();
            const keys = Object.keys(r);
            const cols = [...keys, "row_version", "updated_at", "deleted"];
            const ph = cols.map(_ => "?").join(",");
            const vals = [...keys.map(k => r[k]), rv, now, 0];
            db.prepare(`INSERT INTO ${table} (${cols.join(",")}) VALUES (${ph})`).run(...vals);
        }
    });
    tx();
    const maxv = Number((db.prepare(`SELECT value FROM meta WHERE key='max_row_version'`).get() as any).value);
    notifyChange(table, maxv);
    return true;
};


export const changes = (_: any, { table, since_version, limit = 5000, offset = 0 }: any) => {
    const def = assertKnownTable(table);
    // row_version > since_version 인 행을 최신 상태로 그대로 보내고,
    // deleted 플래그 기준으로 op를 도출한다.
    const rows = db.prepare(`
    SELECT * FROM ${table}
    WHERE row_version > ?
    ORDER BY row_version ASC
    LIMIT ? OFFSET ?
  `).all(since_version, limit, offset);

    const out = rows.map((r: any) => ({
        row: r,
        row_version: r.row_version,
        op: r.deleted ? "delete" : "upsert",
    }));
    return out;
};