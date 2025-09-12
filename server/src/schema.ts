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

type AuditEntry { id: Int!, ts: String!, actor: String!, action: String!, table_name: String, detail: String }

input ColumnDefIn { name: String!, type: String!, notNull: Boolean, default: JSON, check: String }
input UpsertRowInput { id: Int!, base_row_version: Int, data: JSON! }

type Query {
  meta: Meta!
  rows(table: String!, since_version: Int, whereRaw: String, orderBy: String, limit: Int, offset: Int, include_deleted: Boolean): RowResult!
  presence: [Presence!]!
  locks(sheet: String): [Lock!]!
  auditLog(actor: String, action: String, table: String, since: String, until: String, limit: Int, offset: Int): [AuditEntry!]!
}

type Mutation {
  # 스키마
  createTable(table: String!, columns: [ColumnDefIn!]!): Boolean!
  addColumns(table: String!, columns: [ColumnDefIn!]!): Boolean!
  addIndex(table: String!, name: String!, expr: String!, unique: Boolean): Boolean!

  # 데이터
  upsertRows(table: String!, rows: [JSON!]!, actor: String!): RowResult!
  upsertRowsV2(table: String!, rows: [UpsertRowInput!]!, actor: String!): RowResult!
  deleteRows(table: String!, ids: [Int!]!, actor: String!): RowResult!

  # Presence/락
  presenceHeartbeat(nickname: String!, sheet: String, cell: String): Boolean!
  acquireLock(sheet: String!, cell: String!, nickname: String!): Boolean!
  releaseLock(sheet: String!, cell: String!, nickname: String!): Boolean!

  # 복구
  recoverFromExcel(table: String!, rows: [JSON!]!, schema_hash: String!, actor: String!): Boolean!
}
`;
