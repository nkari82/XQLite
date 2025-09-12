// server/src/resolvers/audit.ts
import { db } from "../db";

export type AuditRow = {
    id: number;
    ts: string;
    actor: string;
    action: string;
    table_name?: string | null;
    detail?: string | null;
};

/**
 * 다른 리졸버에서 감사 로그를 남길 때 사용.
 * rows.ts에 이미 직접 INSERT가 있지만, 통일하려면 이 함수를 사용하세요.
 */
export function writeAudit(
    actor: string,
    action: string,
    table_name?: string,
    detail?: unknown
) {
    db.prepare(
        `INSERT INTO audit_log(actor, action, table_name, detail)
     VALUES (?, ?, ?, ?)`
    ).run(actor, action, table_name ?? null, detail == null ? null : JSON.stringify(detail));
}

/**
 * (옵션) 감사 로그 조회 리졸버
 * 사용하려면 GraphQL SDL에 아래를 추가하고 index.ts에서 resolvers.Query.auditLog = queryAudit 로 연결하세요.
 *
 * type AuditEntry { id: Int!, ts: String!, actor: String!, action: String!, table_name: String, detail: String }
 * extend type Query {
 *   auditLog(actor: String, action: String, table: String, since: String, until: String, limit: Int, offset: Int): [AuditEntry!]!
 * }
 */
export const queryAudit = (_: any, args: {
    actor?: string;
    action?: string;
    table?: string;
    since?: string; // ISO8601 e.g., "2025-09-01T00:00:00Z"
    until?: string; // ISO8601
    limit?: number;
    offset?: number;
}): AuditRow[] => {
    const conds: string[] = [];
    const params: any[] = [];

    if (args.actor) { conds.push("actor = ?"); params.push(args.actor); }
    if (args.action) { conds.push("action = ?"); params.push(args.action); }
    if (args.table) { conds.push("table_name = ?"); params.push(args.table); }
    if (args.since) { conds.push("ts >= ?"); params.push(args.since); }
    if (args.until) { conds.push("ts < ?"); params.push(args.until); }

    const where = conds.length ? `WHERE ${conds.join(" AND ")}` : "";
    const limit = Number.isFinite(args.limit) ? Math.max(1, Math.min(1000, args.limit as number)) : 200;
    const offset = Number.isFinite(args.offset) ? Math.max(0, args.offset as number) : 0;

    const sql = `
    SELECT id, ts, actor, action, table_name, detail
    FROM audit_log
    ${where}
    ORDER BY id DESC
    LIMIT ? OFFSET ?`;
    return db.prepare(sql).all(...params, limit, offset) as AuditRow[];
};

/**
 * (옵션) 보관주기 관리용: cutoff 이전 로그 삭제
 * 예: purgeBefore("2025-01-01T00:00:00Z")
 */
export function purgeBefore(cutoffISO8601: string): number {
    const info = db.prepare(`DELETE FROM audit_log WHERE ts < ?`).run(cutoffISO8601);
    return info.changes ?? 0;
}
