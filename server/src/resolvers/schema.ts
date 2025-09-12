import { db } from "../db";

export const createTable = (_: any, { table, columns }: { table: string, columns: any[] }) => {
    const userCols = columns.map(c => `${c.name} ${c.type}`).join(", ");
    const metaCols = `id INTEGER PRIMARY KEY, row_version INTEGER NOT NULL, updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP, deleted INTEGER NOT NULL DEFAULT 0`;
    db.exec(`CREATE TABLE IF NOT EXISTS ${table} (${metaCols}${userCols ? ", " + userCols : ""})`);
    db.exec(`CREATE INDEX IF NOT EXISTS ix_${table}_row_version ON ${table}(row_version)`);
    return true;
};

export const addColumns = (_: any, { table, columns }: { table: string, columns: any[] }) => {
    const stmt = db.transaction(() => {
        for (const c of columns) {
            db.exec(`ALTER TABLE ${table} ADD COLUMN ${c.name} ${c.type}`);
        }
    });
    stmt();
    return true;
};

export const addIndex = (_: any, { table, name, expr, unique }: { table: string, name: string, expr: string, unique?: boolean }) => {
    db.exec(`CREATE ${unique ? "UNIQUE" : ""} INDEX IF NOT EXISTS ${name} ON ${table}(${expr})`);
    return true;
};

