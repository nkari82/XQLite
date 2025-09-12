export const typeDefs = `#graphql
scalar JSON

type Meta { schema_hash: String!, max_row_version: Int! }
type RowResult {
  rows: [JSON!]!
  max_row_version: Int!
  affected: Int!
  conflicts: [JSON!]
  errors: [String!]
}

type Presence { nickname: String!, sheet: String, cell: String, updated_at: String! }
type Lock { sheet: String!, cell: String!, nickname: String!, updated_at: String! }

type Query {
  meta: Meta!
  rows(table: String!, since_version: Int, whereRaw: String, orderBy: String, limit: Int, offset: Int, include_deleted: Boolean): RowResult!
  presence: [Presence!]!
  locks(sheet: String): [Lock!]!
}

type Mutation {
  createTable(table: String!, columns: [JSON!]!): Boolean!
  addColumns(table: String!, columns: [JSON!]!): Boolean!
  addIndex(table: String!, name: String!, expr: String!, unique: Boolean): Boolean!

  upsertRows(table: String!, rows: [JSON!]!, actor: String!): RowResult!
  deleteRows(table: String!, ids: [Int!]!, actor: String!): RowResult!

  presenceHeartbeat(nickname: String!, sheet: String, cell: String): Boolean!
  acquireLock(sheet: String!, cell: String!, nickname: String!): Boolean!
  releaseLock(sheet: String!, cell: String!, nickname: String!): Boolean!

  recoverFromExcel(table: String!, rows: [JSON!]!, schema_hash: String!, actor: String!): Boolean!
}

input UpsertRowInput {
  id: Int!
  base_row_version: Int     # ← Excel이 보냄(없으면 0)
  data: JSON!               # 실제 컬럼 값들
}

extend type Mutation {
  upsertRowsV2(table: String!, rows: [UpsertRowInput!]!, actor: String!): RowResult!
}
`;
