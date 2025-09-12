// server/src/resolvers/meta.ts
import { db } from "../db";

type Meta = { schema_hash: string; max_row_version: number };

function getMetaValue(key: string, fallback = ""): string {
    const row = db.prepare(`SELECT value FROM meta WHERE key=?`).get(key) as { value?: string } | undefined;
    return row?.value ?? fallback;
}

function setMetaValue(key: string, value: string) {
    db.prepare(
        `INSERT INTO meta(key,value) VALUES(?,?)
     ON CONFLICT(key) DO UPDATE SET value=excluded.value`
    ).run(key, value);
}

export const getMeta = (): Meta => {
    const schema_hash = getMetaValue("schema_hash", "");
    const max_row_version = Number(getMetaValue("max_row_version", "0")) || 0;
    return { schema_hash, max_row_version };
};

// (선택) 외부 모듈에서 쓰려면 export
export const metaStore = { get: getMetaValue, set: setMetaValue };
