import client from "prom-client";

export const registry = new client.Registry();
client.collectDefaultMetrics({ register: registry });

export const gqlCounter = new client.Counter({
    name: "xqlite_gql_requests_total",
    help: "GraphQL requests count",
    labelNames: ["op", "type"],
});
registry.registerMetric(gqlCounter);

export const gqlDuration = new client.Histogram({
    name: "xqlite_gql_duration_seconds",
    help: "GraphQL execution time",
    buckets: [0.005, 0.01, 0.02, 0.05, 0.1, 0.2, 0.5, 1, 2],
    labelNames: ["op"],
});
registry.registerMetric(gqlDuration);
