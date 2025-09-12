import { db, presenceTTLSeconds } from "../db";

export const heartbeat = (_: any, { nickname, sheet, cell }: { nickname: string, sheet?: string, cell?: string }) => {
    db.prepare(`
    INSERT INTO presence(nickname, sheet, cell, updated_at)
    VALUES (?, ?, ?, CURRENT_TIMESTAMP)
    ON CONFLICT(nickname) DO UPDATE SET sheet=excluded.sheet, cell=excluded.cell, updated_at=CURRENT_TIMESTAMP
  `).run(nickname, sheet ?? null, cell ?? null);
    return true;
};

export const queryPresence = () => {
    return db.prepare(`
    SELECT nickname, sheet, cell, updated_at
    FROM presence
    WHERE (strftime('%s','now') - strftime('%s',updated_at)) <= ?
  `).all(presenceTTLSeconds);
};

export const acquire = (_: any, { sheet, cell, nickname }: { sheet: string, cell: string, nickname: string }) => {
    try {
        db.prepare(`
      INSERT INTO locks(sheet,cell,nickname,updated_at)
      VALUES (?,?,?,CURRENT_TIMESTAMP)
      ON CONFLICT(sheet,cell) DO UPDATE SET
        nickname=CASE WHEN (strftime('%s','now') - strftime('%s',updated_at)) > ? THEN excluded.nickname ELSE locks.nickname END,
        updated_at=CASE WHEN (strftime('%s','now') - strftime('%s',updated_at)) > ? THEN CURRENT_TIMESTAMP ELSE locks.updated_at END
    `).run(sheet, cell, nickname, presenceTTLSeconds, presenceTTLSeconds);
        const row = db.prepare(`SELECT nickname FROM locks WHERE sheet=? AND cell=?`).get(sheet, cell) as { nickname?: string } | undefined;
        return row?.nickname === nickname;
    } catch { return false; }
};

export const release = (_: any, { sheet, cell, nickname }: { sheet: string, cell: string, nickname: string }) => {
    const row = db.prepare(`SELECT nickname FROM locks WHERE sheet=? AND cell=?`).get(sheet, cell) as { nickname?: string } | undefined;
    if (row?.nickname === nickname) {
        db.prepare(`DELETE FROM locks WHERE sheet=? AND cell=?`).run(sheet, cell);
        return true;
    }
    return false;
};

export const queryLocks = (_: any, { sheet }: { sheet?: string }) => {
    if (sheet) return db.prepare(`SELECT * FROM locks WHERE sheet=?`).all(sheet);
    return db.prepare(`SELECT * FROM locks`).all();
};
