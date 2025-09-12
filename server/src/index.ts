// Apollo Server v5 + Express 5 (최신)
import { ApolloServer } from '@apollo/server';
import { expressMiddleware } from '@as-integrations/express5';
import { ApolloServerPluginDrainHttpServer } from '@apollo/server/plugin/drainHttpServer';

import express, { Request, Response, NextFunction } from 'express';
import http from 'http';
import cors from 'cors';
import { rateLimit } from 'express-rate-limit';

import { typeDefs } from './schema.js';
import * as meta from './resolvers/meta.js';
import * as rows from './resolvers/rows.js';
import * as schemaOps from './resolvers/schema.js';
import * as presence from './resolvers/presence.js';
import * as audit from './resolvers/audit.js';

import { config } from './config.js';
import { logger } from './logger.js';
import { registry, gqlCounter, gqlDuration } from './observability.js';
import { integrityCheck } from './maintenance.js';
import { runMigrations } from "./migrator.js";
// ── GraphQL resolvers
const resolvers = {
    Query: {
        meta: meta.getMeta,
        rows: rows.queryRows,
        changes: rows.changes,          // ← 추가
        presence: presence.queryPresence,
        locks: presence.queryLocks,
        auditLog: audit.queryAudit,
    },
    Mutation: {
        createTable: schemaOps.createTable,
        addColumns: schemaOps.addColumns,
        addIndex: schemaOps.addIndex,
        upsertRows: rows.upsertRows,
        deleteRows: rows.deleteRows,
        presenceHeartbeat: presence.heartbeat,
        acquireLock: presence.acquire,
        releaseLock: presence.release,
        recoverFromExcel: rows.recoverFromExcel,
    },
};


if (process.env.MIGRATE_ON_BOOT !== "0") {
    runMigrations(require("path").resolve(process.cwd(), "migrations"));
}


// ── Express 5 앱 구성
const app = express();
const httpServer = http.createServer(app);

// CORS & Body Parser
app.use(cors({ origin: config.corsOrigin, credentials: false }));
app.use(express.json({ limit: '3mb' }));

// Rate Limit (v8: named export, 옵션명 `limit`)
const limiter = rateLimit({
    windowMs: 60_000,
    limit: config.rateLimitRPM,
    standardHeaders: 'draft-8',
    legacyHeaders: false,
});
app.use(limiter);

// 간단 API KEY 인증 (프로덕션은 JWT 등 권장)
app.use((req: Request, res: Response, next: NextFunction) => {
    if (config.apiKey && req.headers['x-api-key'] !== config.apiKey) {
        res.status(401).json({ error: 'unauthorized' });
        return;
    }
    next();
});

// 헬스체크 & 메트릭
app.get('/health', async (_req, res) => {
    const ok = integrityCheck(process.env.DB_PATH || 'db.sqlite');
    res.json({ ok, integrity: ok });
});

app.get('/metrics', async (_req, res) => {
    res.setHeader('Content-Type', registry.contentType);
    res.end(await registry.metrics());
});

// ── Apollo Server (플러그인 포함)
type Ctx = { actor: string };

const opLogPlugin = {
    async requestDidStart() {
        const endAll = gqlDuration.startTimer();
        return {
            async didResolveOperation(ctx: any) {
                const op = ctx.request.operationName ?? 'anonymous';
                const type = ctx.operation?.operation ?? 'unknown';
                gqlCounter.inc({ op, type: String(type) });
            },
            async willSendResponse() {
                endAll({ op: 'all' });
            },
        };
    },
};

const server = new ApolloServer<Ctx>({
    typeDefs,
    resolvers,
    plugins: [
        opLogPlugin,
        ApolloServerPluginDrainHttpServer({ httpServer }),
    ],
    formatError: (err) => err,
});

// ── 미들웨어 장착 (/graphql)
await server.start();
app.use(
    '/graphql',
    express.json({ limit: '3mb' }),
    expressMiddleware(server, {
        context: async (req: any) => ({
            actor: String((req.headers['x-actor'] as string) ?? 'unknown'),
        }),
    }),
);

// ── 기동
httpServer.listen(config.port, () => {
    logger.info({ port: config.port }, 'XQLite server up');
});
