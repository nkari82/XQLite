-- 메타 컬럼 + 게임 컬럼
CREATE TABLE IF NOT EXISTS items (
  id INTEGER PRIMARY KEY,
  row_version INTEGER NOT NULL,
  updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
  deleted INTEGER NOT NULL DEFAULT 0,

  name TEXT NOT NULL,
  rarity INTEGER NOT NULL DEFAULT 1 CHECK (rarity BETWEEN 1 AND 5),
  required_lv INTEGER NOT NULL DEFAULT 1 CHECK (required_lv >= 1),
  atk_lv50 INTEGER NOT NULL DEFAULT 0 CHECK (atk_lv50 >= 0),

  -- JSON1 필드: 자유 옵션(예: { "tags":["sword","fire"], "vendor":"shop" })
  options TEXT DEFAULT '{}' CHECK (json_valid(options)),

  -- Generated Column: 레벨당 공격력(가볍게)
  atk_per_lv REAL GENERATED ALWAYS AS (CAST(atk_lv50 AS REAL)/50.0) VIRTUAL
);

CREATE INDEX IF NOT EXISTS ix_items_row_version ON items(row_version);
CREATE INDEX IF NOT EXISTS ix_items_sort ON items(rarity DESC, atk_lv50 DESC, id ASC);
CREATE INDEX IF NOT EXISTS ix_items_required_lv ON items(required_lv);

-- 요약 테이블: 총계와 버전 추적
CREATE TABLE IF NOT EXISTS item_stats (
  key TEXT PRIMARY KEY,
  value INTEGER NOT NULL
);
INSERT OR IGNORE INTO item_stats(key,value) VALUES('total',0),('max_row_version',0);

-- 요약 트리거: insert/update/delete마다 자동 유지(삭제 플래그 반영)
CREATE TRIGGER IF NOT EXISTS tg_items_insert AFTER INSERT ON items
BEGIN
  UPDATE item_stats SET value = value + CASE WHEN NEW.deleted=0 THEN 1 ELSE 0 END WHERE key='total';
  UPDATE item_stats SET value = MAX(value, NEW.row_version) WHERE key='max_row_version';
END;

CREATE TRIGGER IF NOT EXISTS tg_items_update AFTER UPDATE ON items
BEGIN
  -- 삭제 플래그 변경 시 총계 조정
  UPDATE item_stats SET value = value + (CASE
     WHEN OLD.deleted=0 AND NEW.deleted=1 THEN -1
     WHEN OLD.deleted=1 AND NEW.deleted=0 THEN +1
     ELSE 0 END)
  WHERE key='total';
  UPDATE item_stats SET value = MAX(value, NEW.row_version) WHERE key='max_row_version';
END;

CREATE TRIGGER IF NOT EXISTS tg_items_delete AFTER DELETE ON items
BEGIN
  UPDATE item_stats SET value = value + CASE WHEN OLD.deleted=0 THEN -1 ELSE 0 END WHERE key='total';
END;

-- 대표 뷰: 목록용(정렬·파생값 포함)
CREATE VIEW IF NOT EXISTS v_item_list AS
SELECT
  id, name, rarity, required_lv, atk_lv50, atk_per_lv, options,
  row_version, updated_at
FROM items
WHERE deleted=0
ORDER BY rarity DESC, atk_lv50 DESC, id ASC;
