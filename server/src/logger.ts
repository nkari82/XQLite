import pino from 'pino';
import { config } from './config';

export const logger = pino({
    level: config.nodeEnv === 'production' ? 'info' : 'debug',
    transport: config.nodeEnv === 'production' ? undefined : { target: 'pino-pretty' },
});
