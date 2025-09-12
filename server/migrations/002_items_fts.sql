-- FTS5: 이름/태그 검색용 (content=items, content_rowid=id)
-- unicode61 토크나이저 사용
CREATE VIRTUAL TABLE IF NOT EXISTS items_fts
USING fts5(name, tags, content='items', content_rowid='id', tokenize='unicode61');

-- 초기 동기화
INSERT INTO items_fts(rowid, name, tags)
SELECT i.id, i.name, COALESCE(json_extract(i.options,'$.tags'), '')
FROM items i
WHERE i.deleted=0;

-- 동기화 트리거
CREATE TRIGGER IF NOT EXISTS tg_items_ai AFTER INSERT ON items BEGIN
  INSERT INTO items_fts(rowid, name, tags)
  SELECT NEW.id, NEW.name, COALESCE(json_extract(NEW.options,'$.tags'), '')
  WHERE NEW.deleted=0;
END;

CREATE TRIGGER IF NOT EXISTS tg_items_au AFTER UPDATE ON items BEGIN
  DELETE FROM items_fts WHERE rowid=OLD.id;
  INSERT INTO items_fts(rowid, name, tags)
  SELECT NEW.id, NEW.name, COALESCE(json_extract(NEW.options,'$.tags'), '')
  WHERE NEW.deleted=0;
END;

CREATE TRIGGER IF NOT EXISTS tg_items_ad AFTER DELETE ON items BEGIN
  DELETE FROM items_fts WHERE rowid=OLD.id;
END;

-- 검색 뷰(랭크 포함)
CREATE VIEW IF NOT EXISTS v_item_search AS
SELECT i.id, i.name, i.rarity, i.required_lv, i.atk_lv50, i.options, i.row_version,
       bm25(items_fts) AS rank
FROM items_fts
JOIN items i ON i.id = items_fts.rowid
WHERE i.deleted=0
ORDER BY rank ASC; -- 낮을수록 일치
