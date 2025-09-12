// src/sse.ts (TS NodeNext/ESM 기준)
import type { Request, Response, NextFunction, RequestHandler } from 'express';
import express from 'express';
import { rateLimit } from 'express-rate-limit';
import { db, RowData } from './db.js';
import { sseConnections } from './observability.js';

function sendEvent(res: Response, ev: 'change' | 'delete', data: unknown, id: number) {
    // SSE 표준 포맷
    res.write(`id: ${id}\n`);
    res.write(`event: ${ev}\n`);
    res.write(`data: ${JSON.stringify(data)}\n\n`);
}

function toSince(req: Request): number {
    const fromHeader = (req.headers['last-event-id'] as string) ?? '';
    const fromQuery = (req.query.since as string) ?? '';
    const n = parseInt(fromQuery || fromHeader || '0', 10);
    return Number.isFinite(n) && n > 0 ? n : 0;
}

function sseHeaders(res: Response) {
    res.setHeader('Content-Type', 'text/event-stream; charset=utf-8');
    res.setHeader('Cache-Control', 'no-cache, no-transform');
    res.setHeader('Connection', 'keep-alive');

    // 재연결 지시 (클라이언트 기본 재시도 간격)
    res.write('retry: 3000\n\n'); // ⬅️ 추가
    // 중간 프록시 버퍼링 방지(가능한 경우)
    // @ts-ignore
    res.flushHeaders?.();
}

function validateIdent(id: string): boolean {
    // 테이블/뷰 이름 화이트리스트(첫 글자 영문/언더스코어, 이후 영문숫자언더스코어)
    return /^[A-Za-z_][A-Za-z0-9_]*$/.test(id);
}

/**
 * SSE 마운트
 * GET /events/rows/:table?since=123
 *  - since 또는 Last-Event-ID 를 읽어 그 이후 row_version만 전송
 *  - 1s 폴링 + 25s 하트비트
 */
export function mountSSE(
    app: express.Express,
    opts?: {
        basePath?: string;               // 기본: '/events'
        middlewares?: RequestHandler[];  // 예: [apiKey, sseLimiter]
        pollMs?: number;                 // 기본: 1000ms
        heartbeatMs?: number;            // 기본: 25000ms
        pageLimit?: number;              // 기본: 256
    }
) {
    const base = opts?.basePath ?? '/events';
    const mws = opts?.middlewares ?? [];
    const pollMs = opts?.pollMs ?? 1000;
    const heartbeatMs = opts?.heartbeatMs ?? 25_000;
    const pageLimit = opts?.pageLimit ?? 256;

    // 재연결 폭주 대비 전용 리밋(원하면 외부에서 주입도 가능)
    const defaultLimiter = rateLimit({ windowMs: 60_000, limit: 60 });
    const chain = mws.length ? mws : [defaultLimiter];

    app.get(`${base}/rows/:table`, ...chain, async (req: Request, res: Response, next: NextFunction) => {
        try {
            const table = String(req.params.table || '');
            if (!validateIdent(table)) {
                res.status(400).json({ error: 'invalid table' });
                return;
            }


            // 안전하게 컬럼 바인딩만 파라미터화; 테이블명은 화이트리스트 통과 후 템플릿
            const selectSql = `SELECT * FROM ${table} WHERE row_version > ? ORDER BY row_version LIMIT ?`;
            const stmt = db.prepare(selectSql);

            let lastSent = toSince(req);

            sseHeaders(res);

            // 첫 페이지 즉시 푸시
            const pushOnce = () => {
                const rows = stmt.all(lastSent, pageLimit) as RowData[];
                for (const r of rows) {
                    const ev: 'change' | 'delete' = !!r.deleted ? 'delete' : 'change';
                    const id = r.row_version;
                    sendEvent(res, ev, { row: r, row_version: id }, id);
                    lastSent = id;
                }
            };

            pushOnce();

            const pollTimer = setInterval(pushOnce, pollMs);
            const heartbeatTimer = setInterval(() => res.write(`: keep-alive ${Date.now()}\n\n`), heartbeatMs);

            sseConnections.inc();      // ⬅️ 연결 증가
            // 연결 종료 처리
            const cleanup = () => {
                clearInterval(pollTimer);
                clearInterval(heartbeatTimer);
                // SSE는 res.end()를 서버가 먼저 호출하지 않는게 일반적이나,
                // 커넥션이 닫히면 Express가 정리한다.
                sseConnections.dec();    // ⬅️ 종료 감소
            };

            req.on('close', cleanup);
            req.on('aborted', cleanup);
        } catch (e) {
            next(e);
        }
    });
}
