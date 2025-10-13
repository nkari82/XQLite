# XQLite Source Review

## Repository Overview
- **GraphQL server (`Server/src/index.ts`)** – Node.js service using `graphql-yoga`, `better-sqlite3`, and file-system access to manage collaborative spreadsheet data, including schema management, synchronization, audit logging, and a pub/sub channel for real-time updates.【F:Server/src/index.ts†L1-L705】
- **Excel add-in (`XQLiteAddIn/*.cs`)** – Excel-DNA based add-in that bootstraps a GraphQL backend client, orchestrates synchronization/collaboration helpers, and persists client configuration (endpoint, API key, timings) to disk or environment-driven overrides.【F:XQLiteAddIn/XqlAddIn.cs†L1-L171】【F:XQLiteAddIn/XqlConfig.cs†L1-L163】

## High-Risk Findings

### Unauthenticated administrative GraphQL surface
The server exposes a single Yoga endpoint without any authentication or API-key validation, yet the Excel client expects to send an API key. Every resolver operates without caller identity checks, meaning any network client can invoke them to read/write the database or manage locks.【F:Server/src/index.ts†L498-L705】【F:XQLiteAddIn/XqlAddIn.cs†L61-L146】【F:XQLiteAddIn/XqlConfig.cs†L15-L135】

### Full database exfiltration helper
The `exportDatabase` query reads the entire SQLite file from disk and returns it base64-encoded. With no access control, this grants unauthenticated callers a complete offline copy of all data, including audit history and tombstones.【F:Server/src/index.ts†L520-L555】

### Dynamic schema manipulation via unsanitized identifiers
Mutations such as `createTable`, `addColumns`, `dropColumns`, and `deleteRow` accept arbitrary table/column identifiers and interpolate them directly into SQL strings. Although double quotes are applied, the strings are not escaped, allowing crafted names with quotes to break out and execute injected SQL. Attackers can create/destroy tables or run arbitrary statements against the shared database.【F:Server/src/index.ts†L110-L167】【F:Server/src/index.ts†L584-L666】

### Error detail leakage
Yoga is initialized with `maskedErrors: false`, causing internal stack traces and detailed error messages to be forwarded to clients. Combined with the lack of authentication, this increases the risk of information disclosure during probing or exploitation attempts.【F:Server/src/index.ts†L691-L705】

### Plain-text client configuration (including API key)
The add-in persists the GraphQL endpoint, API key, and collaboration settings directly to a JSON file inside the user profile without encryption or OS credential APIs. Malware or other users on the machine can harvest the API key and endpoint simply by reading the config file.【F:XQLiteAddIn/XqlConfig.cs†L15-L93】

## Additional Observations
- The `meta` and `tableColumns` queries enumerate schema metadata, which aids reconnaissance for attackers once authenticated endpoints are introduced.【F:Server/src/index.ts†L531-L555】
- Presence, audit, and lock management resolvers (`presence`, `audit_log`, `presenceTouch`, `acquireLock`, etc.) accept arbitrary nicknames/identifiers, so bogus clients can spoof other users or flood audit logs unless server-side identity checks are added.【F:Server/src/index.ts†L498-L666】
- Configuration loading honors environment variables and sidecar files without validation; hardening should include signing or trusted path enforcement if used in shared environments.【F:XQLiteAddIn/XqlAddIn.cs†L128-L169】【F:XQLiteAddIn/XqlConfig.cs†L46-L163】

## Suggested Next Steps
1. Introduce authentication/authorization on the GraphQL layer (e.g., API key headers validated before resolver execution) and enforce TLS.【F:Server/src/index.ts†L498-L705】
2. Sanitize or parameterize dynamic SQL identifiers, or whitelist allowable table/column operations to prevent injection and schema abuse.【F:Server/src/index.ts†L110-L167】【F:Server/src/index.ts†L584-L666】
3. Restrict or remove `exportDatabase`, or require elevated privileges for bulk exports.【F:Server/src/index.ts†L531-L534】
4. Enable `maskedErrors` (or custom error masking) to avoid leaking internal stack traces to clients.【F:Server/src/index.ts†L691-L705】
5. Protect stored API keys—e.g., via Windows DPAPI or credential vaults—and consider multi-user machine implications.【F:XQLiteAddIn/XqlConfig.cs†L15-L93】
