import { db } from "../db.js";
import { z } from "zod";

export type ColumnDef = {
    name: string;
    type: "INTEGER" | "REAL" | "TEXT" | "BLOB" | "BOOLEAN";
    notNull?: boolean;
    default?: any;
    check?: string;
};
export type TableDef = { table: string; columns: ColumnDef[] };

export function getTableDef(table: string): TableDef | null {
    const row = db.prepare(`SELECT value FROM meta WHERE key=?`).get(`schema:${table}`) as any;
    if (!row?.value) return null;
    return JSON.parse(row.value);
}

export function setTableDef(def: TableDef) {
    db.prepare(`
    INSERT INTO meta(key,value) VALUES(?,?)
    ON CONFLICT(key) DO UPDATE SET value=excluded.value
  `).run(`schema:${def.table}`, JSON.stringify(def));
}

export function zodShapeFor(def: TableDef) {
    const shape: Record<string, any> = { id: z.number().int().nonnegative() };
    for (const c of def.columns) {
        let s: any;
        switch (c.type) {
            case "INTEGER": s = z.number().int().nullable(); break;
            case "REAL": s = z.number().nullable(); break;
            case "BOOLEAN": s = z.boolean().nullable(); break;
            case "TEXT": s = z.string().nullable(); break;
            case "BLOB": s = z.any().nullable(); break;
            default: s = z.any();
        }
        if (c.notNull) {
            s = s.nullable().transform((v: any) => {
                if (v == null)
                    throw new Error(`${c.name} NULL`);
                return v;
            });
        }
        shape[c.name] = s;
    }
    return z.object(shape).strict();
}


export function syncRegistryFromDB() {
    const tables = db.prepare(`
    SELECT name FROM sqlite_master
    WHERE type='table' AND name NOT LIKE 'sqlite_%'
      AND name NOT IN ('migrations','audit_log','presence','locks','meta','item_stats')
  `).all().map((r: any) => r.name as string);

    for (const t of tables) {
        const cols = db.prepare(`PRAGMA table_info(${t})`).all() as Array<{ name: string; type: string; notnull: number; dflt_value: any }>;
        // 메타 컬럼 제외
        const userCols = cols.filter(c => !['id', 'row_version', 'updated_at', 'deleted'].includes(c.name))
            .map(c => ({
                name: c.name,
                type: (c.type || "TEXT").toUpperCase() as any,
                notNull: !!c.notnull,
                default: c.dflt_value ?? undefined,
            }));
        const def: TableDef = { table: t, columns: userCols };
        setTableDef(def);
    }
}