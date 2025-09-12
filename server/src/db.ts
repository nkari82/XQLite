import Database = require("better-sqlite3");

export const db = new Database("db.sqlite");
db.pragma("journal_mode = WAL");
db.pragma("synchronous = NORMAL");
db.pragma("busy_timeout = 5000");
// db.pragma("cache_size = -200000"); // 필요 시

// 메타/감사/프레즌스/락/버전 관리 테이블
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

// presence TTL 만료용 뷰(조회 시 필터링)
export const presenceTTLSeconds = 10;


export function nextRowVersion(): number {
    const cur = Number((db.prepare("SELECT value FROM meta WHERE key='max_row_version'").get() as { value: string }).value);
    const nxt = cur + 1;
    db.prepare("UPDATE meta SET value=? WHERE key='max_row_version'").run(String(nxt));
    return nxt;
}
