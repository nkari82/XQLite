import 'dotenv/config';

function req(name: string, def?: string) {
    const v = process.env[name] ?? def;
    if (v == null) throw new Error(`Missing env: ${name}`);
    return v;
}

export const config = {
    port: Number(req('PORT', '4000')),
    nodeEnv: req('NODE_ENV', 'development'),
    corsOrigin: req('CORS_ORIGIN', 'http://localhost'),
    dbPath: req('DB_PATH', 'db.sqlite'),
    apiKey: req('API_KEY', ''), // 간단 API 키(프로덕션은 JWT/MTLS 권장)
    rateLimitRPM: Number(req('RATE_LIMIT_RPM', '600')),
    presenceTTL: Number(req('PRESENCE_TTL', '10')),
};
