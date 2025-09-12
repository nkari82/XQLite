import http from 'http';
import { logger } from './logger.js';

type Stopper = () => Promise<void>;

export function setupShutdown(httpServer: http.Server, stops: Stopper[]) {
    let stopping = false;

    async function shutdown(signal: string) {
        if (stopping) return;
        stopping = true;
        logger.info({ signal }, 'shutdown_begin');

        // 1) Apollo/내부 드레인 먼저
        for (const stop of stops) {
            try { await stop(); } catch (e) { /* log only */ }
        }

        // 2) HTTP 서버 종료
        await new Promise<void>((resolve) => {
            httpServer.close(() => resolve());
            // 강제 타임아웃(10s)
            setTimeout(resolve, 10_000).unref();
        });

        logger.info('shutdown_complete');
        process.exit(0);
    }

    process.on('SIGINT', () => shutdown('SIGINT'));
    process.on('SIGTERM', () => shutdown('SIGTERM'));
}
