// XqlGqlBackend.cs (async-first, 정리 버전)

using GraphQL;
using GraphQL.Client.Http;
using GraphQL.Client.Serializer.Newtonsoft;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reactive.Linq;
using System.Threading;
using System.Threading.Tasks;
using static XQLite.AddIn.IXqlBackend;

namespace XQLite.AddIn
{


    // =======================================================================
    // Backend 인터페이스 (Async 전용)
    // =======================================================================
    internal interface IXqlBackend : IDisposable
    {
        internal enum ConnState { Disconnected, Connecting, Online, Degraded }

        // Sync (데이터 동기화)
        Task<UpsertResult> UpsertCells(IEnumerable<EditCell> cells, CancellationToken ct = default);

        Task<PullResult> PullRows(long since, CancellationToken ct = default);
        void StartSubscription(Action<ServerEvent> onEvent, long since);
        void StopSubscription();

        // Collab (Presence/Lock)
        Task PresenceTouch(string nickname, string? sheet, string? cell, CancellationToken ct = default);
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


        // 연결상태 하트비트(부작용 없음) — 호출할 때마다 상태 업데이트
        Task<long> PingAsync(CancellationToken ct = default);

        Task<List<ColumnInfo>> GetTableColumns(string table, CancellationToken ct = default);

        Task TryDropColumns(string table, IEnumerable<string> names, CancellationToken ct = default);

        // 상태 조회/이벤트
        event Action<ConnState, string?>? StateChanged;
        ConnState State { get; }
        string? StateDetail { get; }
        DateTime LastOkUtc { get; }
    }

    // =======================================================================
    // 실 구현: GraphQL Backend
    // =======================================================================
    internal sealed class XqlGqlBackend : IXqlBackend
    {
        // ── GraphQL 문서 (서버 스키마에 맞춰 사용) ───────────────────────────────

        // 1) Pull: Long → Int
        private const string Q_PULL = @"query($since:Int!){
            rows(since_version:$since){
                max_row_version
                patches { table row_key row_version deleted cells }
            }
        }";

        // 2) Subscription: rowsChanged → events, 변수 제거
        private const string SUB_ROWS = @"subscription{
            events{
                max_row_version
                patches { table row_key row_version deleted cells }
            }
        }";

        private const string MUT_UPSERT_CELLS = @"
            mutation($cells:[CellEditInput!]!){
                upsertCells(cells:$cells){
                max_row_version
                errors
                conflicts { table row_key column message }
            }
        }";

        // 3) PresenceTouch는 동일 (변경 없음)
        private const string MUT_PRESENCE = @"mutation($n:String!,$s:String,$c:String){
            presenceTouch(nickname:$n, sheet:$s, cell:$c){ ok }
        }";

        private const string MUT_ACQUIRE = @"mutation($cell:String!, $by:String!){ acquireLock(cell:$cell, by:$by){ ok } }";
        private const string MUT_RELEASE_BY = @"mutation($by:String!){ releaseLocksBy(by:$by){ ok } }";

        private const string MUT_CREATE_TABLE = @"mutation($table:String!, $key:String!){
          createTable(table:$table, key:$key){ ok }
        }";

        // 4) addColumns: notnull → notNull
        private const string MUT_ADD_COLUMNS = @"mutation($table:String!, $columns:[ColumnDefInput!]!){
            addColumns(table:$table, columns:$columns){ ok }
        }";

        // 5) upsertRows: 서버 반환과 일치하게 selection 축소
        private const string MUT_UPSERT_ROWS = @"mutation ($table:String!,$rows:[JSON!]!){
                upsertRows(table:$table, rows:$rows){
                max_row_version
                errors
            }
        }";

        // 6) meta: JSON 스칼라 → 필드 선택 없이 그대로
        private const string Q_META = @"query{ meta }";

        // 7) audit: Long → Int
        private const string Q_AUDIT = @"query($since:Int){
                audit_log(since_version:$since){
                ts user table row_key column old_value new_value row_version
            }
        }";

        private const string Q_EXPORT_DB = @"query{ exportDatabase }";

        private const string Q_PRESENCE = @"query { presence { nickname sheet cell updated_at } }";

        // ── GQL
        private const string Q_TABLE_COLUMNS = @"query($t:String!){
            tableColumns(table:$t){ name type notnull pk }
        }";

        private const string MUT_DROP_COLUMNS = @"mutation($t:String!,$ns:[String!]!){
            dropColumns(table:$t, names:$ns){ ok }
        }";

        // ── 필드 ─────────────────────────────────────────────────────────────
        private readonly GraphQLHttpClient _http;
        private readonly GraphQLHttpClient _ws;
        private IDisposable? _subscription;
        private int _subRetry = 0;

        private readonly Timer _heartbeat;
        public ConnState State { get; private set; } = ConnState.Connecting;
        public string? StateDetail { get; private set; }
        public DateTime LastOkUtc { get; private set; } = DateTime.MinValue;
        public event Action<ConnState, string?>? StateChanged;

        private const int HB_TTL_MS = 10_000; // 마지막 성공 이후 이 시간 넘도록 성공 없으면 Disconnected

        // ── 생성자 ───────────────────────────────────────────────────────────
        internal XqlGqlBackend(string httpEndpoint, string? apiKey, int heartbeatSec = 3)
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

            _heartbeat = new Timer(async _ => await SafeHeartbeat(), null, Timeout.Infinite, Timeout.Infinite);
            _ = SafeHeartbeat(); // 즉시 1회
            _heartbeat.Change(TimeSpan.FromSeconds(heartbeatSec), TimeSpan.FromSeconds(heartbeatSec));
        }

        public void Dispose()
        {
            try { StopSubscription(); } catch { }
            try { _http.Dispose(); } catch { }
            try { _ws.Dispose(); } catch { }
            try
            {
                _heartbeat.Change(Timeout.Infinite, Timeout.Infinite);
                _heartbeat.Dispose();
            }
            catch { }
        }

        // ── 상태 갱신 ────────────────────────────────────────────────────────
        private void SetState(ConnState st, string? detail = null)
        {
            if (State == st && detail == StateDetail) return;
            State = st;
            StateDetail = detail;
            if (st == ConnState.Online) LastOkUtc = DateTime.UtcNow;
            try { StateChanged?.Invoke(st, detail); } catch { }
        }

        // ⬇️ 연결상태 하트비트 (호출 시 상태 업데이트)
        public async Task<long> PingAsync(CancellationToken ct = default)
        {
            try
            {
                var req = new GraphQLHttpRequest { Query = "query { ping }" };
                var resp = await _http.SendQueryAsync<PingQueryResult>(req, ct).ConfigureAwait(false);
                var now = resp.Data?.ping ?? 0L;
                SetState(ConnState.Online, "ping ok");
                return now;
            }
            catch (Exception ex)
            {
                var now = DateTime.UtcNow;
                if (LastOkUtc == DateTime.MinValue)
                {
                    SetState(ConnState.Connecting, "ping...");
                }
                else
                {
                    var ms = (now - LastOkUtc).TotalMilliseconds;
                    if (ms > HB_TTL_MS) SetState(ConnState.Disconnected, $"ping timeout {(int)(ms / 1000)}s");
                    else SetState(ConnState.Degraded, $"ping fail: {ex.GetType().Name}");
                }
                throw; // 호출측에서 필요 시 무시 가능
            }
        }

        private async Task SafeHeartbeat()
        {
            if (_heartbeat == null)
                return;

            try
            {
                // 연결상태 확인 (상태는 PingAsync 내부에서 갱신됨)
                await PingAsync().ConfigureAwait(false);
            }
            catch { /* 네트워크 일시 오류는 무시 */ }
        }


        // ── Sync ─────────────────────────────────────────────────────────────
        public async Task<PullResult> PullRows(long since, CancellationToken ct = default)
        {
            // 서버 rows(since_version: Int!) 스키마 보호
            var since32 = (int)Math.Min(Math.Max(since, int.MinValue), int.MaxValue);

            var req = new GraphQLRequest { Query = Q_PULL, Variables = new { since = since32 } };
            var resp = await _http.SendQueryAsync<JObject>(req, ct).ConfigureAwait(false);
            return ParsePull(resp.Data);
        }


        public void StartSubscription(Action<ServerEvent> onEvent, long since)
        {
            StopSubscription();
            // 변수 없이 바로 구독
            var req = new GraphQLRequest { Query = SUB_ROWS };
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
        // ⬇️ 프레즌스 갱신 (연결상태와 무관)
        public async Task PresenceTouch(string nickname, string? sheet, string? cell, CancellationToken ct = default)
        {
            var req = new GraphQLHttpRequest
            {
                Query = MUT_PRESENCE,
                Variables = new { n = nickname, s = sheet, c = cell }
            };
            var resp = await _http.SendMutationAsync<PresenceTouchMutation>(req, ct).ConfigureAwait(false);
            if (resp.Errors != null && resp.Errors.Length > 0)
                throw new Exception("presenceTouch failed: " + resp.Errors[0].Message);
        }

        public Task AcquireLock(string cellOrResourceKey, string by, CancellationToken ct = default) =>
            _http.SendMutationAsync<JObject>(new GraphQLRequest { Query = MUT_ACQUIRE, Variables = new { cell = cellOrResourceKey, by } }, ct);

        public Task ReleaseLocksBy(string by, CancellationToken ct = default) =>
            _http.SendMutationAsync<JObject>(new GraphQLRequest { Query = MUT_RELEASE_BY, Variables = new { by } }, ct);

        // ── Backup / Schema ──────────────────────────────────────────────────
        public Task TryCreateTable(string table, string key, CancellationToken ct = default)
            => _http.SendMutationAsync<object>(new GraphQLRequest
            {
                Query = MUT_CREATE_TABLE,
                Variables = new { table, key }
            }, ct);

        public Task TryAddColumns(string table, IEnumerable<ColumnDef> cols, CancellationToken ct = default)
            => _http.SendMutationAsync<object>(new GraphQLRequest
            {
                Query = MUT_ADD_COLUMNS,
                Variables = new
                {
                    table,
                    columns = cols.Select(c => new
                    {
                        name = c.Name,
                        kind = c.Kind,
                        notNull = c.NotNull,
                        check = c.Check
                    }).ToArray()
                }
            }, ct);

        public async Task<JObject?> TryFetchServerMeta(CancellationToken ct = default)
        {
            var resp = await _http.SendQueryAsync<JObject>(new GraphQLRequest { Query = Q_META }, ct).ConfigureAwait(false);
            return resp.Data?["meta"] as JObject; // 서버가 JSON 스칼라를 반환
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
            var req = new GraphQLHttpRequest
            {
                Query = Q_PRESENCE
            };

            var resp = await _http.SendQueryAsync<PresenceResp>(req, ct).ConfigureAwait(false);
            var arr = resp.Data?.presence ?? Array.Empty<PresenceItem>();

            // 혹시라도 역직렬화 중 null 항목이 섞이면 제거
            return arr.Where(p => p != null).ToArray();
        }

        public async Task<bool> UpsertRows(string table, List<Dictionary<string, object?>> rows, CancellationToken ct = default)
        {
            var req = new GraphQLRequest { Query = MUT_UPSERT_ROWS, Variables = new { table, rows } };
            var resp = await _http.SendMutationAsync<JObject>(req, ct).ConfigureAwait(false);
            var parsed = ParseUpsert(resp.Data);
            return parsed.Errors == null || parsed.Errors.Count == 0;
        }

        public async Task<UpsertResult> UpsertCells(IEnumerable<EditCell> cells, CancellationToken ct = default)
        {
            var req = new GraphQLRequest { Query = MUT_UPSERT_CELLS, Variables = new { cells } };
            var resp = await _http.SendMutationAsync<JObject>(req, ct).ConfigureAwait(false);

            var root = resp.Data?["upsertCells"] as JObject
                       ?? throw new Exception("upsertCells: empty response");

            return new UpsertResult
            {
                MaxRowVersion = (long?)root["max_row_version"] ?? 0,
                Errors = root["errors"] is JArray ea ? ea.Select(x => x?.ToString() ?? "").ToList() : null,
                Conflicts = root["conflicts"] is JArray ca ? ca.ToObject<List<Conflict>>() : null
            };
        }

        // ── 호출
        public async Task<List<ColumnInfo>> GetTableColumns(string table, CancellationToken ct = default)
        {
            var req = new GraphQLRequest { Query = Q_TABLE_COLUMNS, Variables = new { t = table } };
            var resp = await _http.SendQueryAsync<JObject>(req, ct).ConfigureAwait(false);
            var arr = resp.Data?["tableColumns"] as JArray;
            return arr?.ToObject<List<ColumnInfo>>() ?? new List<ColumnInfo>();
        }

        public Task TryDropColumns(string table, IEnumerable<string> names, CancellationToken ct = default)
        {
            var list = names?.Distinct(StringComparer.Ordinal) ?? Enumerable.Empty<string>();
            return _http.SendMutationAsync<object>(new GraphQLRequest
            {
                Query = MUT_DROP_COLUMNS,
                Variables = new { t = table, ns = list.ToArray() }
            }, ct);
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

            // ✅ 서버는 events
            var root = data["events"] as JObject;
            // 혹시 구독 구현에서 배열로 올 수도 있으니 방어
            if (root == null && data["events"] is JArray arr && arr.Count > 0)
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

    internal sealed class PingQueryResult { public long ping { get; set; } }
    internal sealed class PresenceTouchMutation { public OkResult? presenceTouch { get; set; } }
    internal sealed class OkResult { public bool ok { get; set; } }

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
    internal sealed class PresenceItem
    {
        public string? nickname { get; set; }
        public string? sheet { get; set; }
        public string? cell { get; set; }
        public long? updated_at { get; set; } // ← ms epoch
    }

    internal sealed class PresenceResp
    {
        public PresenceItem[]? presence { get; set; }
    }

    internal sealed class UpsertResp
    {
        public UpsertRowsPayload? upsertRows { get; set; }
    }

    internal sealed class UpsertRowsPayload
    {
        public int affected { get; set; }
        public GqlErr[]? errors { get; set; }
        public long max_row_version { get; set; }
    }

    internal sealed class GqlErr
    {
        public string? code { get; set; }
        public string? message { get; set; }
    }


    // Parser 결과 DTO

    public sealed class ColumnInfo
    {
        public string name { get; set; } = "";
        public string? type { get; set; }
        public bool notnull { get; set; }
        public bool pk { get; set; }
    }

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
