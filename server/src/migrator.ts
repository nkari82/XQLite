import fs from "fs";
import path from "path";
import crypto from "crypto";
import { db } from "./db.js";
import { logger } from "./logger.js";
import { config } from "./config.js";

type MigRow = { id: number; name: string; checksum: string; applied_at: string };

function sha256(s: string) {
    return crypto.createHash("sha256").update(s, "utf8").digest("hex");
}

function ensureSystemTables() {
    db.exec(`
    CREATE TABLE IF NOT EXISTS migrations (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      name TEXT NOT NULL UNIQUE,
      checksum TEXT NOT NULL,
      applied_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
    );
  `);
}

export function computeSchemaHash(): string {
    // 스키마 해시(오브젝트 이름 순서로 정렬된 SQL을 연결하여 해시)
    const rows = db.prepare(`
    SELECT type, name, sql
    FROM sqlite_master
    WHERE type IN ('table','index','view','trigger')
      AND name NOT LIKE 'sqlite_%'
    ORDER BY type, name
  `).all() as Array<{ type: string; name: string; sql: string | null }>;
    const bundle = rows.map(r => `-- ${r.type}:${r.name}\n${r.sql ?? ""}\n`).join("\n");
    return sha256(bundle);
}

function assertExtAvailable() {
    // FTS5/JSON1 가용성 점검(없으면 경고만)
    const opts = db.prepare(`PRAGMA compile_options;`).all().map((r: any) => Object.values(r)[0] as string);
    const hasFTS5 = opts.some(o => /FTS5/i.test(o));
    const hasJSON1 = opts.some(o => /JSON1/i.test(o));
    if (!hasFTS5) logger.warn("FTS5 extension not detected in compile options. (Most prebuilt better-sqlite3 include it.)");
    if (!hasJSON1) logger.warn("JSON1 extension not detected. (Modern SQLite usually has it.)");
}

export function runMigrations(migrationsDir = path.resolve(process.cwd(), "migrations")) {
    ensureSystemTables();
    assertExtAvailable();

    if (!fs.existsSync(migrationsDir)) {
        logger.info({ migrationsDir }, "no migrations directory, skipping");
        return;
    }

    const files = fs.readdirSync(migrationsDir)
        .filter(f => f.endsWith(".sql"))
        .sort((a, b) => a.localeCompare(b));

    const applied = new Map<string, MigRow>();
    for (const r of db.prepare(`SELECT * FROM migrations ORDER BY id`).all() as MigRow[]) {
        applied.set(r.name, r);
    }

    const tx = db.transaction(() => {
        for (const f of files) {
            const full = path.join(migrationsDir, f);
            const sql = fs.readFileSync(full, "utf8");
            const sum = sha256(sql);

            if (applied.has(f)) {
                const row = applied.get(f)!;
                if (row.checksum !== sum) {
                    throw new Error(`Migration checksum mismatch for ${f}. Previously applied: ${row.checksum}, now: ${sum}`);
                }
                continue; // already applied
            }

            db.exec(sql);
            db.prepare(`INSERT INTO migrations(name, checksum) VALUES (?,?)`).run(f, sum);
            logger.info({ migration: f }, "applied");
        }

        // 마지막에 스키마 해시 업데이트
        const schemaHash = computeSchemaHash();
        db.prepare(`
      INSERT INTO meta(key,value) VALUES('schema_hash',?)
      ON CONFLICT(key) DO UPDATE SET value=excluded.value
    `).run(schemaHash);
    });

    tx();
    logger.info("migrations complete");
}
