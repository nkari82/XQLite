import Database = require('better-sqlite3');
import { logger } from './logger.js';

export function snapshot(dbPath: string, outPath: string) {
    const src = new Database(dbPath, { readonly: true });
    const dst = new Database(outPath);
    // Online backup API
    // @ts-ignore
    src.backup(outPath).then(() => {
        logger.info({ outPath }, 'snapshot done');
        src.close(); dst.close();
    }).catch((e: any) => logger.error({ e }, 'snapshot failed'));
}

export function integrityCheck(dbPath: string) {
    const db = new Database(dbPath, { readonly: true });
    const row = db.prepare(`PRAGMA integrity_check;`).get() as any;
    db.close();
    return row?.integrity_check === 'ok';
}
