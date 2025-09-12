import { TableDef } from "../resolvers/registry.js";

// "colA DESC, colB ASC" 형태만 허용, 컬럼 화이트리스트 검사
export function sanitizeOrderBy(input: string | undefined, def: TableDef): string {
    if (!input) return "";
    const allowed = new Set(def.columns.map(c => c.name).concat(["id", "row_version", "updated_at", "deleted"]));
    const parts = input.split(",").map(p => p.trim()).filter(Boolean);
    const out: string[] = [];
    for (const p of parts) {
        const m = /^([A-Za-z_][A-Za-z0-9_\.]*)(\s+(ASC|DESC))?$/i.exec(p);
        if (!m) continue;
        const col = m[1];
        const dir = (m[3] || "ASC").toUpperCase();
        if (!allowed.has(col)) continue;
        out.push(`${col} ${dir}`);
    }
    return out.length ? `ORDER BY ${out.join(", ")}` : "";
}
