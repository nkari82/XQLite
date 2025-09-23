// XqlGqlBackend.cs
using GraphQL;
using GraphQL.Client.Http;
using GraphQL.Client.Serializer.Newtonsoft;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;

namespace XQLite.AddIn
{
    
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
            new Conflict { Kind = "system", Message = $"[{where}] {msg}" };
    }

    internal sealed class ColumnDef
    {
        public string Name = "";
        public string Kind = "text";   // int/real/text/bool/json/date
        public bool NotNull = false;
        public string? Check;
    }

    internal interface IXqlBackend : IDisposable
    {
        // sync
        UpsertResult UpsertCells(IReadOnlyList<EditCell> cells);
        PullResult PullRows(long since);
        void StartSubscription(Action<ServerEvent> onEvent, long since);
        void StopSubscription();

        // collab
        void PresenceHeartbeat(string nickname, string? cell);
        void AcquireLock(string cell, string by);
        void ReleaseLocksBy(string by);

        // backup/schema
        void TryCreateTable(string table, string key);
        void TryAddColumns(string table, IEnumerable<ColumnDef> cols);
        JObject? TryFetchServerMeta();
        JArray? TryFetchAuditLog(long? since = null);
        byte[]? TryExportDatabase();
    }

    internal sealed class XqlGqlBackend : IXqlBackend
    {
        // === GQL 문서(프로젝트 스키마 맞춰 사용) ===
        const string MUT_UPSERT = @"mutation($cells:[CellEditInput!]!){
          upsertCells(cells:$cells){ max_row_version errors conflicts { table row_key column message server_version local_version } }
        }";
        const string Q_PULL = @"query($since:Long!){
          rows(since_version:$since){ max_row_version patches { table row_key row_version deleted cells } }
        }";
        const string SUB_ROWS = @"subscription($since:Long){
          rowsChanged(since_version:$since){ max_row_version patches { table row_key row_version deleted cells } }
        }";
        const string MUT_HEARTBEAT = @"mutation($nickname:String!, $cell:String){
          presenceHeartbeat(nickname:$nickname, cell:$cell) { ok }
        }";
        const string MUT_ACQUIRE = @"mutation($cell:String!, $by:String!){
          acquireLock(cell:$cell, by:$by) { ok }
        }";
        const string MUT_RELEASE_BY = @"mutation($by:String!){
          releaseLocksBy(by:$by) { ok }
        }";
        const string MUT_CREATE_TABLE = @"mutation($table:String!, $key:String!){
          createTable(table:$table, key:$key){ ok }
        }";
        const string MUT_ADD_COLUMNS = @"mutation($table:String!, $columns:[ColumnDefInput!]!){
          addColumns(table:$table, columns:$columns){ ok }
        }";
        const string Q_META = @"query{ meta{ schema_hash max_row_version tables{ name cols{ name kind notnull } } } }";
        const string Q_AUDIT = @"query($since:Long){
          audit_log(since_version:$since){ ts user table row_key column old_value new_value row_version }
        }";
        const string Q_EXPORT_DB = @"query{ exportDatabase }";

        private readonly GraphQLHttpClient _http;
        private readonly GraphQLHttpClient _ws;
        private IDisposable? _subscription;

        internal XqlGqlBackend(string httpEndpoint, string? apiKey)
        {
            var httpUri = new Uri(httpEndpoint);
            var wsUri = new UriBuilder(httpUri) { Scheme = httpUri.Scheme == "https" ? "wss" : "ws" }.Uri;

            _http = new GraphQLHttpClient(new GraphQLHttpClientOptions { EndPoint = httpUri }, new NewtonsoftJsonSerializer());
            _ws = new GraphQLHttpClient(new GraphQLHttpClientOptions { EndPoint = wsUri, UseWebSocketForQueriesAndMutations = false }, new NewtonsoftJsonSerializer());
            if (!string.IsNullOrWhiteSpace(apiKey))
            {
                _http.HttpClient.DefaultRequestHeaders.Add("x-api-key", apiKey);
                _ws.HttpClient.DefaultRequestHeaders.Add("x-api-key", apiKey);
            }
        }

        public void Dispose() 
        { 
            try 
            { 
                StopSubscription(); 
            } 
            catch { } 
            
            try 
            { 
                _http.Dispose(); 
            } 
            catch { } 
            
            try 
            { 
                _ws.Dispose(); 
            } catch { } 
        }

        // --- Sync ---
        public UpsertResult UpsertCells(IReadOnlyList<EditCell> cells)
        {
            var req = new GraphQLRequest
            {
                Query = MUT_UPSERT,
                Variables = new
                {
                    cells = cells.Select(c => new { table = c.Table, row_key = c.RowKey, column = c.Column, value = c.Value }).ToArray()
                }
            };
            var resp = _http.SendQueryAsync<JObject>(req).GetAwaiter().GetResult();
            return ParseUpsert(resp.Data);
        }

        public PullResult PullRows(long since)
        {
            var req = new GraphQLRequest { Query = Q_PULL, Variables = new { since } };
            var resp = _http.SendQueryAsync<JObject>(req).GetAwaiter().GetResult();
            return ParsePull(resp.Data);
        }

        public void StartSubscription(Action<ServerEvent> onEvent, long since)
        {
            StopSubscription();
            var req = new GraphQLRequest { Query = SUB_ROWS, Variables = new { since } };
            var observable = _ws.CreateSubscriptionStream<JObject>(req);
            _subscription = observable.Subscribe(
                p => { try { onEvent(ParseSub(p.Data)); } catch { } },
                _ => { try { StopSubscription(); new Timer(_ => StartSubscription(onEvent, since), null, 2000, Timeout.Infinite); } catch { } },
                () => { new Timer(_ => StartSubscription(onEvent, since), null, 2000, Timeout.Infinite); });
        }

        public void StopSubscription() { try { _subscription?.Dispose(); } catch { } _subscription = null; }

        // --- Collab ---
        public void PresenceHeartbeat(string nickname, string? cell)
            => _http.SendMutationAsync<JObject>(new GraphQLRequest { Query = MUT_HEARTBEAT, Variables = new { nickname, cell } }).GetAwaiter().GetResult();
        public void AcquireLock(string cell, string by)
            => _http.SendMutationAsync<JObject>(new GraphQLRequest { Query = MUT_ACQUIRE, Variables = new { cell, by } }).GetAwaiter().GetResult();
        public void ReleaseLocksBy(string by)
            => _http.SendMutationAsync<JObject>(new GraphQLRequest { Query = MUT_RELEASE_BY, Variables = new { by } }).GetAwaiter().GetResult();

        // --- Backup/Schema ---
        public void TryCreateTable(string table, string key)
            => _http.SendMutationAsync<JObject>(new GraphQLRequest { Query = MUT_CREATE_TABLE, Variables = new { table, key } }).GetAwaiter().GetResult();
        public void TryAddColumns(string table, IEnumerable<ColumnDef> cols)
            => _http.SendMutationAsync<JObject>(new GraphQLRequest
            {
                Query = MUT_ADD_COLUMNS,
                Variables = new { table, columns = cols }
            }).GetAwaiter().GetResult();
        public JObject? TryFetchServerMeta()
            => _http.SendQueryAsync<JObject>(new GraphQLRequest { Query = Q_META }).GetAwaiter().GetResult()?.Data?["meta"] as JObject;
        public JArray? TryFetchAuditLog(long? since = null)
            => _http.SendQueryAsync<JObject>(new GraphQLRequest { Query = Q_AUDIT, Variables = new { since } }).GetAwaiter().GetResult()?.Data?["audit_log"] as JArray;
        public byte[]? TryExportDatabase()
        {
            var d = _http.SendQueryAsync<JObject>(new GraphQLRequest { Query = Q_EXPORT_DB }).GetAwaiter().GetResult()?.Data?["exportDatabase"]?.ToString();
            if (string.IsNullOrWhiteSpace(d)) return null;
            try { return Convert.FromBase64String(d); } catch { return null; }
        }

        // --- Parsers (기존 XqlSync 파서 이동) ---
        private static UpsertResult ParseUpsert(JObject? data) 
        { 
            /* XqlSync.ParseUpsert 내용 이동 */ 
            return UpsertResult.From(data); 
        }

        private static PullResult ParsePull(JObject? data) 
        { 
            /* XqlSync.ParsePull  내용 이동 */ 
            return PullResult.From(data); 
        }

        private static ServerEvent ParseSub(JObject? data) 
        { 
            /* XqlSync.ParseSub   내용 이동 */ 
            return ServerEvent.From(data); 
        }
    }

    // DTO들: XqlSync의 형식 재사용(팩토리 메서드 추가)
    internal sealed class UpsertResult
    {
        public long MaxRowVersion; 
        public List<string>? Errors; 
        public List<Conflict>? Conflicts;

        public static UpsertResult From(JObject? data) 
        {
            /* 기존 로직 이식 */
            var res = new UpsertResult { MaxRowVersion = 0, Errors = new List<string>(), Conflicts = new List<Conflict>() };
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
    }
    internal sealed class PullResult
    {
        public long MaxRowVersion; 
        public List<RowPatch>? Patches;

        public static PullResult From(JObject? data) 
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
#pragma warning disable CS8604 // 가능한 null 참조 인수입니다.
                    var rp = new RowPatch
                    {
                        Table = p["table"]?.ToString() ?? "",
                        RowKey = p["row_key"]?.ToObject<object>() ?? 0,
                        RowVersion = (long?)p["row_version"] ?? 0,
                        Deleted = p["deleted"]?.Type == JTokenType.Boolean && (bool)p["deleted"],
                        Cells = new Dictionary<string, object?>(StringComparer.Ordinal)
                    };
#pragma warning restore CS8604 // 가능한 null 참조 인수입니다.

                    if (p["cells"] is JObject cc)
                        foreach (var prop in cc.Properties())
                            rp.Cells[prop.Name] = prop.Value.Type == JTokenType.Null ? null : prop.Value.ToObject<object?>();

                    res.Patches!.Add(rp);
                }
            }
            return res;
        }
    }
    internal sealed class ServerEvent
    {
        public long MaxRowVersion; 
        public List<RowPatch>? Patches;
        public static ServerEvent From(JObject? data) 
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
#pragma warning disable CS8604 // 가능한 null 참조 인수입니다.
                    var rp = new RowPatch
                    {
                        Table = p["table"]?.ToString() ?? "",
                        RowKey = p["row_key"]?.ToObject<object>() ?? 0,
                        RowVersion = (long?)p["row_version"] ?? 0,
                        Deleted = p["deleted"]?.Type == JTokenType.Boolean && (bool)p["deleted"],
                        Cells = new Dictionary<string, object?>(StringComparer.Ordinal)
                    };
#pragma warning restore CS8604 // 가능한 null 참조 인수입니다.

                    if (p["cells"] is JObject cc)
                        foreach (var prop in cc.Properties())
                            rp.Cells[prop.Name] = prop.Value.Type == JTokenType.Null ? null : prop.Value.ToObject<object?>();

                    ev.Patches!.Add(rp);
                }
            }
            return ev;
        }
    }
}
