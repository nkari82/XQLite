// XqlGqlBackend.cs (async-first, 정리 버전)

using GraphQL;
using GraphQL.Client.Http;
using GraphQL.Client.Serializer.Newtonsoft;
using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reactive.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace XQLite.AddIn
{
    // =======================================================================
    // Backend 인터페이스 (Async 전용)
    // =======================================================================
    internal interface IXqlBackend : IDisposable
    {
        // Sync (데이터 동기화)
        Task<UpsertResult> UpsertCells(IReadOnlyList<EditCell> cells, CancellationToken ct = default);
        Task<PullResult> PullRows(long since, CancellationToken ct = default);
        void StartSubscription(Action<ServerEvent> onEvent, long since);
        void StopSubscription();

        // Collab (Presence/Lock)
        Task PresenceHeartbeat(string nickname, string? cell, CancellationToken ct = default);
        Task AcquireLock(string cellOrResourceKey, string by, CancellationToken ct = default);
        Task ReleaseLocksBy(string by, CancellationToken ct = default);

        // Backup / Schema
        Task TryCreateTable(string table, string key, CancellationToken ct = default);
        Task TryAddColumns(string table, IEnumerable<ColumnDef> cols, CancellationToken ct = default);
        Task<JObject?> TryFetchServerMeta(CancellationToken ct = default);
        Task<JArray?> TryFetchAuditLog(long? since = null, CancellationToken ct = default);
        Task<byte[]?> TryExportDatabase(CancellationToken ct = default);

        // Presence
        Task<PresenceItem[]?> FetchPresence(CancellationToken ct = default);

        // Recover
        Task<bool> UpsertRows(string table, List<Dictionary<string, object?>> rows, CancellationToken ct = default);
    }

    // =======================================================================
    // 실 구현: GraphQL Backend
    // =======================================================================
    internal sealed class XqlGqlBackend : IXqlBackend
    {
        // ── GraphQL 문서 (서버 스키마에 맞춰 사용) ───────────────────────────────
        private const string MUT_UPSERT = @"mutation($cells:[CellEditInput!]!){
          upsertCells(cells:$cells){
            max_row_version
            errors
            conflicts { table row_key column message server_version local_version }
          }
        }";

        private const string Q_PULL = @"query($since:Long!){
          rows(since_version:$since){
            max_row_version
            patches { table row_key row_version deleted cells }
          }
        }";

        private const string SUB_ROWS = @"subscription($since:Long){
          rowsChanged(since_version:$since){
            max_row_version
            patches { table row_key row_version deleted cells }
          }
        }";

        private const string MUT_HEARTBEAT = @"mutation($nickname:String!, $cell:String){
          presenceHeartbeat(nickname:$nickname, cell:$cell){ ok }
        }";

        private const string MUT_ACQUIRE = @"mutation($cell:String!, $by:String!){ acquireLock(cell:$cell, by:$by){ ok } }";
        private const string MUT_RELEASE_BY = @"mutation($by:String!){ releaseLocksBy(by:$by){ ok } }";

        private const string MUT_CREATE_TABLE = @"mutation($table:String!, $key:String!){
          createTable(table:$table, key:$key){ ok }
        }";

        private const string MUT_ADD_COLUMNS = @"mutation($table:String!, $columns:[ColumnDefInput!]!){
          addColumns(table:$table, columns:$columns){ ok }
        }";

        private const string MUT_UPSERT_ROWS = @"mutation ($table:String!,$rows:[JSON!]!){
              upsertRows(table:$table, rows:$rows){
                affected
                errors { code message }
                max_row_version
              }
            }";

        private const string Q_META = @"query{
          meta{ schema_hash max_row_version tables{ name cols{ name kind notnull } } }
        }";

        private const string Q_AUDIT = @"query($since:Long){
          audit_log(since_version:$since){
            ts user table row_key column old_value new_value row_version
          }
        }";

        private const string Q_EXPORT_DB = @"query{ exportDatabase }";

        private const string Q_PRESENCE = @"query { presence { nickname sheet cell updated_at } }";

        // ── 필드 ─────────────────────────────────────────────────────────────
        private readonly GraphQLHttpClient _http;
        private readonly GraphQLHttpClient _ws;
        private IDisposable? _subscription;
        private int _subRetry = 0;

        // ── 생성자 ───────────────────────────────────────────────────────────
        internal XqlGqlBackend(string httpEndpoint, string? apiKey)
        {
            var httpUri = new Uri(httpEndpoint);
            var wsUri = new UriBuilder(httpUri) { Scheme = httpUri.Scheme == "https" ? "wss" : "ws" }.Uri;

            _http = new GraphQLHttpClient(
                new GraphQLHttpClientOptions { EndPoint = httpUri },
                new NewtonsoftJsonSerializer());

            _ws = new GraphQLHttpClient(
                new GraphQLHttpClientOptions { EndPoint = wsUri, UseWebSocketForQueriesAndMutations = false },
                new NewtonsoftJsonSerializer());

            if (!string.IsNullOrWhiteSpace(apiKey))
            {
                _http.HttpClient.DefaultRequestHeaders.Add("x-api-key", apiKey);
                _ws.HttpClient.DefaultRequestHeaders.Add("x-api-key", apiKey);
            }
        }

        public void Dispose()
        {
            try { StopSubscription(); } catch { }
            try { _http.Dispose(); } catch { }
            try { _ws.Dispose(); } catch { }
        }

        // ── Sync ─────────────────────────────────────────────────────────────
        public async Task<UpsertResult> UpsertCells(IReadOnlyList<EditCell> cells, CancellationToken ct = default)
        {
            var req = new GraphQLRequest
            {
                Query = MUT_UPSERT,
                Variables = new
                {
                    cells = cells.Select(c => new
                    {
                        table = c.Table,
                        row_key = c.RowKey,
                        column = c.Column,
                        value = c.Value
                    }).ToArray()
                }
            };

            // 뮤테이션은 SendMutationAsync 사용
            var resp = await _http.SendMutationAsync<JObject>(req, ct).ConfigureAwait(false);
            return ParseUpsert(resp.Data);
        }

        public async Task<PullResult> PullRows(long since, CancellationToken ct = default)
        {
            var req = new GraphQLRequest { Query = Q_PULL, Variables = new { since } };
            var resp = await _http.SendQueryAsync<JObject>(req, ct).ConfigureAwait(false);
            return ParsePull(resp.Data);
        }

        public void StartSubscription(Action<ServerEvent> onEvent, long since)
        {
            StopSubscription();
            var req = new GraphQLRequest { Query = SUB_ROWS, Variables = new { since } };
            var observable = _ws.CreateSubscriptionStream<JObject>(req);

            var sub = observable.Subscribe(
            onNext: p => { _subRetry = 0; try { onEvent(ParseSub(p.Data)); } catch { } },
            onError: _ => Resubscribe(onEvent, since),
            onCompleted: () => Resubscribe(onEvent, since));
            Interlocked.Exchange(ref _subscription, sub)?.Dispose();
        }

        private void Resubscribe(Action<ServerEvent> onEvent, long since)
        {
            StopSubscription();
            var delayMs = (int)Math.Min(30_000, 500 * Math.Pow(2, Math.Min(_subRetry++, 10)));
            _ = new Timer(_ => StartSubscription(onEvent, since), null, delayMs, Timeout.Infinite);
        }

        public void StopSubscription()
        {
            try { _subscription?.Dispose(); } catch { }
            _subscription = null;
        }

        // ── Collab ───────────────────────────────────────────────────────────
        public Task PresenceHeartbeat(string nickname, string? cell, CancellationToken ct = default) =>
            _http.SendMutationAsync<JObject>(new GraphQLRequest { Query = MUT_HEARTBEAT, Variables = new { nickname, cell } }, ct);

        public Task AcquireLock(string cellOrResourceKey, string by, CancellationToken ct = default) =>
            _http.SendMutationAsync<JObject>(new GraphQLRequest { Query = MUT_ACQUIRE, Variables = new { cell = cellOrResourceKey, by } }, ct);

        public Task ReleaseLocksBy(string by, CancellationToken ct = default) =>
            _http.SendMutationAsync<JObject>(new GraphQLRequest { Query = MUT_RELEASE_BY, Variables = new { by } }, ct);

        // ── Backup / Schema ──────────────────────────────────────────────────
        public Task TryCreateTable(string table, string key, CancellationToken ct = default) =>
            _http.SendMutationAsync<JObject>(new GraphQLRequest { Query = MUT_CREATE_TABLE, Variables = new { table, key } }, ct);

        public Task TryAddColumns(string table, IEnumerable<ColumnDef> cols, CancellationToken ct = default) =>
            _http.SendMutationAsync<JObject>(new GraphQLRequest
            {
                Query = MUT_ADD_COLUMNS,
                Variables = new
                {
                    table,
                    columns = cols.Select(c => new
                    {
                        name = c.Name,
                        kind = c.Kind,
                        notnull = c.NotNull,
                        check = c.Check
                    }).ToArray()
                }
            }, ct);

        public async Task<JObject?> TryFetchServerMeta(CancellationToken ct = default)
        {
            var resp = await _http.SendQueryAsync<JObject>(new GraphQLRequest { Query = Q_META }, ct).ConfigureAwait(false);
            return resp.Data?["meta"] as JObject;
        }

        public async Task<JArray?> TryFetchAuditLog(long? since = null, CancellationToken ct = default)
        {
            var resp = await _http.SendQueryAsync<JObject>(new GraphQLRequest { Query = Q_AUDIT, Variables = new { since } }, ct).ConfigureAwait(false);
            return resp.Data?["audit_log"] as JArray;
        }

        public async Task<byte[]?> TryExportDatabase(CancellationToken ct = default)
        {
            var resp = await _http.SendQueryAsync<JObject>(new GraphQLRequest { Query = Q_EXPORT_DB }, ct).ConfigureAwait(false);
            var s = (string?)resp.Data?["exportDatabase"];
            if (string.IsNullOrWhiteSpace(s)) return null;
            try { return Convert.FromBase64String(s); }
            catch { return null; }
        }

        // ── Presence / Recover ───────────────────────────────────────────────
        public async Task<PresenceItem[]?> FetchPresence(CancellationToken ct = default)
        {
            try
            {
                var req = new GraphQLRequest { Query = Q_PRESENCE };
                var resp = await _http.SendQueryAsync<PresenceResp>(req, ct).ConfigureAwait(false);
                return resp?.Data?.presence;
            }
            catch { return null; }
        }

        public async Task<bool> UpsertRows(string table, List<Dictionary<string, object?>> rows, CancellationToken ct = default)
        {
            var req = new GraphQLRequest { Query = MUT_UPSERT_ROWS, Variables = new { table, rows } };
            var resp = await _http.SendMutationAsync<UpsertResp>(req, ct).ConfigureAwait(false);
            var data = resp?.Data?.upsertRows;
            return data != null && (data.errors == null || data.errors.Length == 0);
        }

        // ── Parser bridge ────────────────────────────────────────────────────
        private static UpsertResult ParseUpsert(JObject? data)
        {
            var res = new UpsertResult
            {
                MaxRowVersion = 0,
                Errors = new List<string>(),
                Conflicts = new List<Conflict>()
            };

            if (data == null) return res;

            var root = data["upsertCells"] ?? data["upsertRows"];
            if (root is not JObject u) return res;

            res.MaxRowVersion = (long?)u["max_row_version"] ?? 0;

            if (u["errors"] is JArray errs)
                foreach (var e in errs) res.Errors!.Add(e?.ToString() ?? "");

            if (u["conflicts"] is JArray cts)
            {
                foreach (var c in cts.OfType<JObject>())
                {
                    res.Conflicts!.Add(new Conflict
                    {
                        Kind = "conflict",
                        Table = c["table"]?.ToString() ?? "",
                        RowKey = c["row_key"]?.ToObject<object>(),
                        Column = c["column"]?.ToString(),
                        Message = c["message"]?.ToString() ?? "",
                        ServerVersion = (long?)c["server_version"],
                        LocalVersion = (long?)c["local_version"],
                    });
                }
            }
            return res;
        }

        private static PullResult ParsePull(JObject? data)
        {
            var res = new PullResult { MaxRowVersion = 0, Patches = new List<RowPatch>() };
            if (data == null) return res;

            var root = data["rows"] as JObject;
            if (root == null) return res;

            res.MaxRowVersion = (long?)root["max_row_version"] ?? 0;

            if (root["patches"] is JArray pts)
            {
                foreach (var p in pts.OfType<JObject>())
                {
                    var rp = new RowPatch
                    {
                        Table = p["table"]?.ToString() ?? "",
                        RowKey = p["row_key"]?.ToObject<object>() ?? 0,
                        RowVersion = (long?)p["row_version"] ?? 0,
                        Deleted = p["deleted"]?.Type == JTokenType.Boolean && (bool)p["deleted"]!,
                        Cells = new Dictionary<string, object?>(StringComparer.Ordinal)
                    };

                    if (p["cells"] is JObject cc)
                        foreach (var prop in cc.Properties())
                            rp.Cells[prop.Name] = prop.Value.Type == JTokenType.Null ? null : prop.Value.ToObject<object?>();

                    res.Patches!.Add(rp);
                }
            }
            return res;
        }

        private static ServerEvent ParseSub(JObject? data)
        {
            var ev = new ServerEvent { MaxRowVersion = 0, Patches = new List<RowPatch>() };
            if (data == null) return ev;

            var root = data["rowsChanged"] as JObject;
            if (root == null && data["rowsChanged"] is JArray arr && arr.Count > 0)
                root = arr[0] as JObject;
            if (root == null) return ev;

            ev.MaxRowVersion = (long?)root["max_row_version"] ?? 0;

            if (root["patches"] is JArray pts)
            {
                foreach (var p in pts.OfType<JObject>())
                {
                    var rp = new RowPatch
                    {
                        Table = p["table"]?.ToString() ?? "",
                        RowKey = p["row_key"]?.ToObject<object>() ?? 0,
                        RowVersion = (long?)p["row_version"] ?? 0,
                        Deleted = p["deleted"]?.Type == JTokenType.Boolean && (bool)p["deleted"]!,
                        Cells = new Dictionary<string, object?>(StringComparer.Ordinal)
                    };

                    if (p["cells"] is JObject cc)
                        foreach (var prop in cc.Properties())
                            rp.Cells[prop.Name] = prop.Value.Type == JTokenType.Null ? null : prop.Value.ToObject<object?>();

                    ev.Patches!.Add(rp);
                }
            }
            return ev;
        }
    }

    // =======================================================================
    // DTOs
    // =======================================================================

    internal readonly record struct EditCell(string Table, object RowKey, string Column, object? Value);

    internal sealed class RowPatch
    {
        public string Table { get; set; } = "";
        public object RowKey { get; set; } = default!;
        public Dictionary<string, object?> Cells { get; set; } = new(StringComparer.Ordinal);
        public long RowVersion { get; set; }
        public bool Deleted { get; set; }
    }

    internal sealed class Conflict
    {
        public string Kind { get; set; } = "conflict";
        public string Table { get; set; } = "";
        public object? RowKey { get; set; }
        public string? Column { get; set; }
        public string Message { get; set; } = "";
        public long? ServerVersion { get; set; }
        public long? LocalVersion { get; set; }

        public static Conflict System(string where, string msg) =>
            new()
            { Kind = "system", Message = $"[{where}] {msg}" };
    }

    internal sealed class ColumnDef
    {
        public string Name = "";
        public string Kind = "text";   // int/real/text/bool/json/date
        public bool NotNull = false;
        public string? Check;
    }

    // GraphQL 응답 DTO
    internal sealed class PresenceResp { public PresenceItem[]? presence { get; set; } }
    internal sealed class PresenceItem { public string? nickname { get; set; } public string? sheet { get; set; } public string? cell { get; set; } public string? updated_at { get; set; } }
    internal sealed class UpsertResp { public UpsertRowsPayload? upsertRows { get; set; } }
    internal sealed class UpsertRowsPayload { public int affected { get; set; } public GqlErr[]? errors { get; set; } public long max_row_version { get; set; } }
    internal sealed class GqlErr { public string? code { get; set; } public string? message { get; set; } }

    // Parser 결과 DTO
    internal sealed class UpsertResult
    {
        public long MaxRowVersion;
        public List<string>? Errors;
        public List<Conflict>? Conflicts;
    }

    internal sealed class PullResult
    {
        public long MaxRowVersion;
        public List<RowPatch>? Patches;
    }

    internal sealed class ServerEvent
    {
        public long MaxRowVersion;
        public List<RowPatch>? Patches;
    }
}
