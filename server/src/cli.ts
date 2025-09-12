#!/usr/bin/env node
import { hideBin } from "yargs/helpers";
import yargs from "yargs";
import { config } from "./config.js";
import { logger } from "./logger.js";
import { integrityCheck, snapshot } from "./maintenance.js";
import { db } from "./db.js";

yargs(hideBin(process.argv))
    .scriptName("xqlite")
    .command("snapshot <out>", "DB 스냅샷 파일 생성", (y) => y.positional("out", { type: "string", describe: "출력 파일 경로" }), (argv) => {
        const out = String(argv.out);
        snapshot(config.dbPath, out);
    })
    .command("integrity", "PRAGMA integrity_check", () => { }, () => {
        const ok = integrityCheck(config.dbPath);
        logger.info({ ok }, "integrity_check");
        if (!ok) process.exit(2);
    })
    .command("dump [table]", "테이블 덤프(JSON Lines)", (y) => y.positional("table", { type: "string" }), (argv) => {
        const t = argv.table as string | undefined;
        if (!t) {
            const list = db.prepare(`SELECT name FROM sqlite_master WHERE type='table' ORDER BY name`).all().map((r: any) => r.name);
            console.log(JSON.stringify({ tables: list }));
            return;
        }
        const it = db.prepare(`SELECT * FROM ${t}`).iterate();
        for (const row of it) console.log(JSON.stringify(row));
    })
    .demandCommand(1)
    .strict()
    .help()
    .parse();
