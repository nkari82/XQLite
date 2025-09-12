// Apollo Server v5 + Express 5 (최신)
import { ApolloServer } from '@apollo/server';
import { expressMiddleware } from '@as-integrations/express5';
import { ApolloServerPluginDrainHttpServer } from '@apollo/server/plugin/drainHttpServer';
import { ApolloServerPluginLandingPageLocalDefault } from '@apollo/server/plugin/landingPage/default'; // ← Sandbox 랜딩페이지

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
import { syncRegistryFromDB } from './resolvers/registry.js'

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

    syncRegistryFromDB();
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
    introspection: true,
    plugins: [
        opLogPlugin,
        ApolloServerPluginDrainHttpServer({ httpServer }),
        ApolloServerPluginLandingPageLocalDefault({ embed: true }) // ← 브라우저 IDE
    ],
    formatError: (err) => err,
});


// production에선 다시 채워야 한다.
// API_KEY=dev-secret-change-me

// 인스펙션 쿼리 판별 도우미
function isIntrospection(body: any) {
    const q = typeof body?.query === 'string' ? body.query : '';
    return q.includes('__schema') || q.includes('__type');
}

// ── 미들웨어 장착 (/graphql)
await server.start();

// CORS 등은 그대로 두고…
app.use('/graphql',
    express.json({ limit: '3mb' }),                    // ⬅️ 먼저 body 파싱
    (req, res, next) => {
        // 1) IDE 로딩(landing page)은 GET → 항상 허용
        if (req.method === 'GET') return next();

        // 2) 개발 모드에서 인스펙션 POST는 허용 (헤더 넣기 전 초기 인스펙션 허용)
        if (process.env.ENABLE_IDE && isIntrospection((req as any).body)) return next();

        // 3) 그 외에는 API Key 검사
        const apiKey = req.headers['x-api-key'];
        if (process.env.API_KEY && apiKey !== process.env.API_KEY) {
            return res.status(401).json({ error: 'unauthorized' });
        }
        next();
    },
    expressMiddleware(server, {
        context: async ({ req }) => ({ actor: String((req.headers['x-actor'] as string) ?? 'unknown') }),
    })
);

// ── 기동
httpServer.listen(config.port, () => {
    logger.info({ port: config.port }, 'XQLite server up');
});
