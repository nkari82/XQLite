import { db } from '../db';
import { z } from 'zod';

export type ColumnDef = { name: string; type: 'INTEGER' | 'REAL' | 'TEXT' | 'BLOB' | 'BOOLEAN'; notNull?: boolean; default?: any; check?: string };
export type TableDef = { table: string; columns: ColumnDef[] };

export function getTableDef(table: string): TableDef | null {
    const row = db.prepare(`SELECT value FROM meta WHERE key=?`).get(`schema:${table}`) as any;
    if (!row?.value) return null;
    return JSON.parse(row.value);
}
export function setTableDef(def: TableDef) {
    db.prepare(`INSERT INTO meta(key,value) VALUES(?,?)
    ON CONFLICT(key) DO UPDATE SET value=excluded.value`).run(`schema:${def.table}`, JSON.stringify(def));
}

export function zodShapeFor(def: TableDef) {
    const shape: Record<string, any> = { id: z.number().int().nonnegative() };
    for (const c of def.columns) {
        let s: any;
        switch (c.type) {
            case 'INTEGER': s = z.number().int().nullable(); break;
            case 'REAL': s = z.number().nullable(); break;
            case 'BOOLEAN': s = z.boolean().nullable(); break;
            case 'TEXT': s = z.string().nullable(); break;
            case 'BLOB': s = z.any().nullable(); break;
            default: s = z.any();
        }
        if (c.notNull) s = s.nullable().transform((v: null) => { if (v == null) throw new Error(`${c.name} NULL`); return v; });
        shape[c.name] = s;
    }
    return z.object(shape).strict();
}
