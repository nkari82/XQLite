import { ApolloServer } from "apollo-server";
import { typeDefs } from "./schema";
import * as meta from "./resolvers/meta";
import * as rows from "./resolvers/rows";
import * as schemaOps from "./resolvers/schema";
import * as presence from "./resolvers/presence";
import * as audit from "./resolvers/audit";

const resolvers = {
    Query: {
        meta: meta.getMeta,
        rows: rows.queryRows,
        presence: presence.queryPresence,
        locks: presence.queryLocks,
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

new ApolloServer({ typeDefs, resolvers }).listen({ port: 4000 }).then(({ url }) => {
    console.log(`XQLite server ready at ${url}`);
});
