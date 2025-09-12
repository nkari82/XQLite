import Database from "better-sqlite3";
import { config } from "./config.js";

export const db = new Database(config.dbPath);
db.pragma("journal_mode = WAL");
db.pragma("synchronous = NORMAL");
db.pragma("busy_timeout = 5000");
// db.pragma("cache_size = -200000"); // 필요 시

// 메타/감사/Presence/락
db.exec(`
CREATE TABLE IF NOT EXISTS meta (
  key TEXT PRIMARY KEY,
  value TEXT NOT NULL
);
INSERT OR IGNORE INTO meta(key,value) VALUES
  ('schema_hash',''), ('max_row_version','0');

CREATE TABLE IF NOT EXISTS audit_log (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  ts DATETIME DEFAULT CURRENT_TIMESTAMP,
  actor TEXT NOT NULL,
  action TEXT NOT NULL,
  table_name TEXT,
  detail TEXT
);

CREATE TABLE IF NOT EXISTS presence (
  nickname TEXT PRIMARY KEY,
  sheet TEXT,
  cell TEXT,
  updated_at DATETIME DEFAULT CURRENT_TIMESTAMP
);

CREATE TABLE IF NOT EXISTS locks (
  sheet TEXT NOT NULL,
  cell TEXT NOT NULL,
  nickname TEXT NOT NULL,
  updated_at DATETIME DEFAULT CURRENT_TIMESTAMP,
  PRIMARY KEY(sheet, cell)
);
`);

export function nextRowVersion(): number {
  const cur = Number((db.prepare("SELECT value FROM meta WHERE key='max_row_version'").get() as { value: string }).value);
  const nxt = cur + 1;
  db.prepare("UPDATE meta SET value=? WHERE key='max_row_version'").run(String(nxt));
  return nxt;
}

export const presenceTTLSeconds = config.presenceTTL;

export type RowData = Record<string, any> & { row_version: number, deleted?: boolean | number | null };