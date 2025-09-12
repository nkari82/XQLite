import { db } from "../db.js";
import { setTableDef, TableDef, ColumnDef } from "./registry.js";

function normType(t: string) {
    const u = t.toUpperCase();
    if (u.includes("INT")) return "INTEGER";
    if (u.includes("REAL") || u.includes("FLOA") || u.includes("DOUB")) return "REAL";
    if (u.includes("BLOB")) return "BLOB";
    if (u.includes("BOOL")) return "BOOLEAN";
    return "TEXT";
}

function metaCols() {
    return `id INTEGER PRIMARY KEY, row_version INTEGER NOT NULL, updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP, deleted INTEGER NOT NULL DEFAULT 0`;
}

export const createTable = (_: any, { table, columns }: { table: string; columns: any[] }) => {
    const userColsDDL = (columns ?? []).map((c) => {
        const ty = normType(c.type);
        const nn = c.notNull ? " NOT NULL" : "";
        const def = c.default != null ? ` DEFAULT ${JSON.stringify(c.default)}` : "";
        const chk = c.check ? ` CHECK(${c.check})` : "";
        return `${c.name} ${ty}${nn}${def}${chk}`;
    }).join(", ");
    db.exec(`CREATE TABLE IF NOT EXISTS ${table} (${metaCols()}${userColsDDL ? ", " + userColsDDL : ""})`);
    db.exec(`CREATE INDEX IF NOT EXISTS ix_${table}_row_version ON ${table}(row_version)`);

    // 레지스트리 기록
    const cols: ColumnDef[] = (columns ?? []).map((c) => ({
        name: c.name,
        type: normType(c.type),
        notNull: !!c.notNull,
        default: c.default,
        check: c.check,
    }));
    setTableDef({ table, columns: cols } as TableDef);
    return true;
};

export const addColumns = (_: any, { table, columns }: { table: string; columns: any[] }) => {
    const tx = db.transaction(() => {
        for (const c of columns ?? []) {
            db.exec(`ALTER TABLE ${table} ADD COLUMN ${c.name} ${normType(c.type)}`);
        }
    });
    tx();

    // 레지스트리 갱신
    const defRow = db.prepare(`SELECT value FROM meta WHERE key=?`).get(`schema:${table}`) as any;
    const def = defRow?.value ? JSON.parse(defRow.value) as TableDef : { table, columns: [] };
    for (const c of columns ?? []) {
        if (!def.columns.find((x) => x.name === c.name)) {
            def.columns.push({ name: c.name, type: normType(c.type) });
        }
    }
    setTableDef(def);
    return true;
};

export const addIndex = (_: any, { table, name, expr, unique }: { table: string; name: string; expr: string; unique?: boolean }) => {
    db.exec(`CREATE ${unique ? "UNIQUE" : ""} INDEX IF NOT EXISTS ${name} ON ${table}(${expr})`);
    return true;
};
