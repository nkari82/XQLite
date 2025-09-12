import { ApolloServer } from 'apollo-server-express'
import { typeDefs } from './schema';
import * as meta from './resolvers/meta';
import * as rows from './resolvers/rows';
import * as schemaOps from './resolvers/schema';
import * as presence from './resolvers/presence';
import * as audit from './resolvers/audit';
import { config } from './config';
import { logger } from './logger';
import rateLimit from 'express-rate-limit';
import express from 'express';
import cors from 'cors';
import http from 'http';

const resolvers = {
    Query: {
        meta: meta.getMeta,
        rows: rows.queryRows,
        presence: presence.queryPresence,
        locks: presence.queryLocks,
        // (옵션) 감사 로그 공개 시
        // auditLog: audit.queryAudit,
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
    }
};

// ── Express 래핑으로 보안/레이트리밋/CORS
const app = express();
app.use(cors({ origin: config.corsOrigin, credentials: false }));
app.use(express.json({ limit: '3mb' }));
app.use(rateLimit({ windowMs: 60_000, max: config.rateLimitRPM }));

// 간단 API KEY 인증 미들웨어(프로덕션은 JWT 등으로 대체)
app.use((req: any, res: any, next: any) => {
    if (config.apiKey && req.headers['x-api-key'] !== config.apiKey) {
        res.status(401).json({ error: 'unauthorized' }); return;
    }
    next();
});

const server = new ApolloServer({
    typeDefs, resolvers,
    context: ({ req }) => ({
        actor: req.headers['x-actor'] ?? 'unknown', // Excel 닉네임 전달
    }),
    formatError: (err) => {
        logger.error({ err }, 'GraphQLError');
        return err;
    }
});

(async () => {
    await server.start();
    //server.applyMiddleware({ app, path: '/' });

    const httpServer = http.createServer(app);
    httpServer.listen(config.port, () => {
        logger.info({ port: config.port }, 'XQLite server up');
    });
})();
