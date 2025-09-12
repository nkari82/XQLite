import { Registry, Counter, Gauge, Histogram, collectDefaultMetrics } from 'prom-client';

export const registry = new Registry();
collectDefaultMetrics({ register: registry });

export const gqlCounter = new Counter({
    name: "xqlite_gql_requests_total",
    help: "GraphQL requests count",
    labelNames: ["op", "type"],
});
registry.registerMetric(gqlCounter);

export const gqlDuration = new Histogram({
    name: "xqlite_gql_duration_seconds",
    help: "GraphQL execution time",
    buckets: [0.005, 0.01, 0.02, 0.05, 0.1, 0.2, 0.5, 1, 2],
    labelNames: ["op"],
});
registry.registerMetric(gqlDuration);

export const sseConnections = new Gauge({
    name: 'sse_connections',
    help: 'current SSE open connections',
});
registry.registerMetric(sseConnections);