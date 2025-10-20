// XqlGqlBackend.cs (async-first, 정리 + 컬럼 Rename/Alter 확장)

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
        Task TryDropColumns(string table, IEnumerable<string> names, CancellationToken ct = default);

        // 스키마 변경 고수준 API
        Task TryRenameColumns(string table, IEnumerable<RenameDef> renames, CancellationToken ct = default);
        Task TryAlterColumns(string table, IEnumerable<AlterDef> alters, CancellationToken ct = default);

        Task<JObject?> TryFetchServerMeta(CancellationToken ct = default);
        Task<JArray?> TryFetchAuditLog(long? since = null, CancellationToken ct = default);
        Task<byte[]?> TryExportDatabase(CancellationToken ct = default);

        // Presence
        Task<PresenceItem[]?> FetchPresence(CancellationToken ct = default);

        // Recover
        Task<bool> UpsertRows(string table, List<Dictionary<string, object?>> rows, CancellationToken ct = default);

        // 연결상태 하트비트(부작용 없음) — 호출할 때마다 상태 업데이트
        Task<long> Ping(CancellationToken ct = default);

        Task<List<ColumnInfo>> GetTableColumns(string table, CancellationToken ct = default);

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
        // ── GraphQL 문서 (서버 스키마에 맞춰 사용) ───────────────────────────

        // 1) Pull: Long → Int
        private const string Q_PULL = @"query($since:Int!){
            rows(since_version:$since){
                max_row_version
                patches { table row_key row_version deleted cells }
            }
        }";

        // 1-2) 테이블 스냅샷(JSON 배열) — selection 없이 호출
        private const string Q_ROWS_SNAPSHOT = @"
        query($t:String!){
          rowsSnapshot(table:$t)
        }";

        // 2) Subscription: rowsChanged → events
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

        // 3) PresenceTouch
        private const string MUT_PRESENCE = @"mutation($n:String!,$s:String,$c:String){
            presenceTouch(nickname:$n, sheet:$s, cell:$c){ ok }
        }";

        private const string MUT_ACQUIRE = @"mutation($cell:String!, $by:String!){ acquireLock(cell:$cell, by:$by){ ok } }";
        private const string MUT_RELEASE_BY = @"mutation($by:String!){ releaseLocksBy(by:$by){ ok } }";

        private const string MUT_CREATE_TABLE = @"mutation($table:String!, $key:String!){
          createTable(table:$table, key:$key){ ok }
        }";

        // 4) addColumns: 서버 ColumnDefInput {name,type,notNull,check}
        private const string MUT_ADD_COLUMNS = @"mutation($table:String!, $columns:[ColumnDefInput!]!){
            addColumns(table:$table, columns:$columns){ ok }
        }";

        // 5) upsertRows
        private const string MUT_UPSERT_ROWS = @"mutation ($table:String!,$rows:[JSON!]!){
            upsertRows(table:$table, rows:$rows){
                max_row_version
                errors
            }
        }";

        // 6) meta (JSON)
        private const string Q_META = @"query{ meta }";

        // 7) audit: Long → Int
        private const string Q_AUDIT = @"query($since:Int){
            audit_log(since_version:$since){
                ts user table row_key column old_value new_value row_version
            }
        }";

        private const string Q_EXPORT_DB = @"query{ exportDatabase }";
        private const string Q_PRESENCE = @"query { presence { nickname sheet cell updated_at } }";

        // ── GQL: 테이블 메타/드랍/리네임/알터 ──────────────────────────────
        private const string Q_TABLE_COLUMNS = @"query($t:String!){
            tableColumns(table:$t){ name type notnull pk }
        }";

        private const string MUT_DROP_COLUMNS = @"mutation($t:String!,$ns:[String!]!){
            dropColumns(table:$t, names:$ns){ ok }
        }";

        private const string MUT_RENAME_COLUMNS = @"mutation($t:String!,$rs:[RenameDefInput!]!){
            renameColumns(table:$t, renames:$rs){ ok }
        }";

        private const string MUT_ALTER_COLUMNS = @"mutation($t:String!,$alters:[AlterDefInput!]!){
            alterColumns(table:$t, alters:$alters){ ok }
        }";

        // ── 필드 ─────────────────────────────────────────────────────────────
        private readonly GraphQLHttpClient _http;
        private readonly GraphQLHttpClient _ws;
        private IDisposable? _subscription;
        private int _subRetry = 0;

        private readonly Timer _heartbeat;
        private Timer? _resubTimer;
        public ConnState State { get; private set; } = ConnState.Connecting;
        public string? StateDetail { get; private set; }
        public DateTime LastOkUtc { get; private set; } = DateTime.MinValue;
        public event Action<ConnState, string?>? StateChanged;

        private const int HB_TTL_MS = 10_000; // 마지막 성공 이후 이 시간 넘도록 성공 없으면 Disconnected

        // ── 생성자 ───────────────────────────────────────────────────────────
        internal XqlGqlBackend(string httpEndpoint, string? apiKey, string? project = null, int heartbeatSec = 3)
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

            // project 헤더는 비어있으면 "default"
            var prj = string.IsNullOrWhiteSpace(project) ? "default" : project!.Trim();
            const string H = "x-project";
            try
            {
                if (_http.HttpClient.DefaultRequestHeaders.Contains(H))
                    _http.HttpClient.DefaultRequestHeaders.Remove(H);
                if (_ws.HttpClient.DefaultRequestHeaders.Contains(H))
                    _ws.HttpClient.DefaultRequestHeaders.Remove(H);

                _http.HttpClient.DefaultRequestHeaders.Add(H, prj);
                _ws.HttpClient.DefaultRequestHeaders.Add(H, prj);
            }
            catch { /* ignore */ }

            _heartbeat = new Timer(async _ => await SafeHeartbeat().ConfigureAwait(false), null, Timeout.Infinite, Timeout.Infinite);
            _ = SafeHeartbeat(); // 즉시 1회
            _heartbeat.Change(TimeSpan.FromSeconds(heartbeatSec), TimeSpan.FromSeconds(heartbeatSec));
        }

        public void Dispose()
        {
            try { StopSubscription(); } catch { }
            try { _resubTimer?.Dispose(); _resubTimer = null; } catch { }
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

        // 연결상태 하트비트
        public async Task<long> Ping(CancellationToken ct = default)
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
                throw;
            }
        }

        private async Task SafeHeartbeat()
        {
            try { await Ping().ConfigureAwait(false); } catch { /* ignore transient */ }
        }

        // ── Sync ─────────────────────────────────────────────────────────────
        public async Task<PullResult> PullRows(long since, CancellationToken ct = default)
        {
            var since32 = (int)Math.Min(Math.Max(since, int.MinValue), int.MaxValue);
            var req = new GraphQLRequest { Query = Q_PULL, Variables = new { since = since32 } };
            var resp = await _http.SendQueryAsync<JObject>(req, ct).ConfigureAwait(false);
            return ParsePull(resp.Data);
        }

        public void StartSubscription(Action<ServerEvent> onEvent, long since)
        {
            StopSubscription();
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
            _resubTimer?.Dispose();
            _resubTimer = new System.Threading.Timer(_ =>
            {
                try { StartSubscription(onEvent, since); }
                finally { _resubTimer?.Dispose(); _resubTimer = null; }
            }, null, delayMs, Timeout.Infinite);
        }

        public void StopSubscription()
        {
            try { _subscription?.Dispose(); } catch { }
            _subscription = null;
        }

        // ── Collab ───────────────────────────────────────────────────────────
        public async Task PresenceTouch(string nickname, string? sheet, string? cell, CancellationToken ct = default)
        {
            var req = new GraphQLRequest { Query = MUT_PRESENCE, Variables = new { n = nickname, s = sheet, c = cell } };
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
                    columns = cols.Where(c => !string.IsNullOrWhiteSpace(c.Name))
                                   .Select(c => new
                                   {
                                       name = c.Name,
                                       type = MapType(c.Kind),   // 서버 ColumnDefInput.type
                                       notNull = c.NotNull,
                                       check = c.Check
                                   }).ToArray()
                }
            }, ct);

        public Task TryDropColumns(string table, IEnumerable<string> names, CancellationToken ct = default)
        {
            var list = names?.Where(n => !string.IsNullOrWhiteSpace(n))
                             .Distinct(StringComparer.OrdinalIgnoreCase)
                             .ToArray() ?? Array.Empty<string>();

            if (list.Length == 0) return Task.CompletedTask;

            return _http.SendMutationAsync<object>(new GraphQLRequest
            {
                Query = MUT_DROP_COLUMNS,
                Variables = new { t = table, ns = list }
            }, ct);
        }

        // 컬럼 이름 변경(복수)
        public Task TryRenameColumns(string table, IEnumerable<RenameDef> renames, CancellationToken ct = default)
        {
            var pairs = renames?.Where(r => !string.IsNullOrWhiteSpace(r.From) && !string.IsNullOrWhiteSpace(r.To) &&
                                            !r.From.Equals(r.To, StringComparison.OrdinalIgnoreCase))
                                .Select(r => new { from = r.From, to = r.To })
                                .ToArray() ?? Array.Empty<object>();

            if (pairs.Length == 0) return Task.CompletedTask;

            return _http.SendMutationAsync<object>(new GraphQLRequest
            {
                Query = MUT_RENAME_COLUMNS,
                Variables = new { t = table, rs = pairs }
            }, ct);
        }

        // 컬럼 타입/NOT NULL/CHECK 변경(복수)
        public Task TryAlterColumns(string table, IEnumerable<AlterDef> alters, CancellationToken ct = default)
        {
            var list = alters?.Where(a => !string.IsNullOrWhiteSpace(a.Name))
                              .Select(a => new
                              {
                                  name = a.Name,
                                  toType = a.ToType != null ? MapType(a.ToType) : null,
                                  toNotNull = a.ToNotNull,
                                  toCheck = a.ToCheck
                              })
                              .ToArray() ?? Array.Empty<object>();

            if (list.Length == 0) return Task.CompletedTask;

            return _http.SendMutationAsync<object>(new GraphQLRequest
            {
                Query = MUT_ALTER_COLUMNS,
                Variables = new { t = table, alters = list }
            }, ct);
        }

        public async Task<JObject?> TryFetchServerMeta(CancellationToken ct = default)
        {
            try
            {
                var resp = await _http.SendQueryAsync<JObject>(new GraphQLRequest { Query = Q_META }, ct).ConfigureAwait(false);
                var raw = resp.Data?["meta"];
                if (raw is JObject jo) return jo;

                if (raw is JValue jv)
                {
                    var s = jv.Type == JTokenType.String ? (string?)jv.Value : jv.ToString(Newtonsoft.Json.Formatting.None);
                    if (!string.IsNullOrWhiteSpace(s))
                    {
                        try { return JObject.Parse(s!); } catch { }
                    }
                }
                return null;
            }
            catch { return null; }
        }

        public async Task<JArray?> TryFetchAuditLog(long? since = null, CancellationToken ct = default)
        {
            int? s = null;
            if (since.HasValue)
            {
                var clamped = Math.Max(int.MinValue, Math.Min(int.MaxValue, since.Value));
                s = unchecked((int)clamped);
            }
            var resp = await _http.SendQueryAsync<JObject>(new GraphQLRequest { Query = Q_AUDIT, Variables = new { since = s } }, ct).ConfigureAwait(false);
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

        // Presence / Recover
        public async Task<PresenceItem[]?> FetchPresence(CancellationToken ct = default)
        {
            var req = new GraphQLRequest { Query = Q_PRESENCE };
            var resp = await _http.SendQueryAsync<PresenceResp>(req, ct).ConfigureAwait(false);
            var arr = resp.Data?.presence ?? Array.Empty<PresenceItem>();
            return arr.Where(p => p != null).ToArray();
        }

        public async Task<bool> UpsertRows(string table, List<Dictionary<string, object?>> rows, CancellationToken ct = default)
        {
            var req = new GraphQLRequest { Query = MUT_UPSERT_ROWS, Variables = new { table, rows } };
            var resp = await _http.SendMutationAsync<JObject>(req, ct).ConfigureAwait(false);
            var parsed = ParseUpsert(resp.Data);
            return parsed.Errors == null || parsed.Errors.Count == 0;
        }

        // 테이블 스냅샷(JSON 배열) → RowPatch 리스트 변환
        public async Task<List<RowPatch>> FetchRowsSnapshot(string table, CancellationToken ct = default)
        {
            try
            {
                var req = new GraphQLRequest { Query = Q_ROWS_SNAPSHOT, Variables = new { t = table } };
                var resp = await _http.SendQueryAsync<JObject>(req, ct).ConfigureAwait(false);
                var arr = resp.Data?["rowsSnapshot"] as JArray;
                var list = new List<RowPatch>();
                if (arr == null) return list;

                foreach (var it in arr.OfType<JObject>())
                {
                    var rk = it["id"]?.ToString() ?? it["row_key"]?.ToString() ?? "";
                    var cells = new Dictionary<string, object?>(StringComparer.Ordinal);
                    foreach (var p in it.Properties())
                    {
                        var pn = p.Name;
                        if (pn.Equals("row_version", StringComparison.OrdinalIgnoreCase)) continue;
                        if (pn.Equals("updated_at", StringComparison.OrdinalIgnoreCase)) continue;
                        if (pn.Equals("deleted", StringComparison.OrdinalIgnoreCase)) continue;
                        if (pn.Equals("id", StringComparison.OrdinalIgnoreCase)) continue;
                        cells[pn] = (p.Value as JValue)?.Value;
                    }
                    list.Add(new RowPatch { Table = table, RowKey = rk, Deleted = false, Cells = cells });
                }
                return list;
            }
            catch { return new List<RowPatch>(); }
        }

        public async Task<UpsertResult> UpsertCells(IEnumerable<EditCell> cells, CancellationToken ct = default)
        {
            var payload = cells.Select(c => new
            {
                table = c.Table,
                row_key = c.RowKey,
                column = c.Column,
                value = c.Value
            }).ToArray();

            var req = new GraphQLRequest { Query = MUT_UPSERT_CELLS, Variables = new { cells = payload } };
            var resp = await _http.SendMutationAsync<JObject>(req, ct).ConfigureAwait(false);

            var root = resp.Data?["upsertCells"] as JObject;
            if (root == null) throw new Exception("upsertCells: empty response");

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
                        Deleted = ParseDeleted(p["deleted"]),
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

            var root = data["events"] as JObject;
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
                        Deleted = ParseDeleted(p["deleted"]),
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

        // Int(0/1) 또는 Boolean 모두 허용
        private static bool ParseDeleted(JToken? tok)
        {
            if (tok == null || tok.Type == JTokenType.Null) return false;
            if (tok.Type == JTokenType.Boolean) return (bool)tok!;
            if (tok.Type == JTokenType.Integer) return ((long)tok!) != 0;
            if (tok.Type == JTokenType.String && long.TryParse((string)tok!, out var n)) return n != 0;
            return false;
        }

        // ── 공통: 엑셀 메타 → 서버 type 매핑 ────────────────────────────────
        private static string MapType(string kind)
        {
            if (string.IsNullOrWhiteSpace(kind)) return "text";
            switch (kind.Trim().ToLowerInvariant())
            {
                case "int":
                case "integer": return "integer";
                case "real":
                case "float":
                case "double": return "real";
                case "bool":
                case "boolean": return "bool";
                case "json": return "json";
                case "date": return "integer"; // epoch ms 저장
                case "text":
                case "string":
                default: return "text";
            }
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
            new() { Kind = "system", Message = $"[{where}] {msg}" };
    }

    // 추가/드랍 입력 DTO (입력 전용)
    internal sealed class ColumnDef
    {
        public string Name = "";
        public string Kind = "text";   // integer/real/text/bool/json/date
        public bool NotNull = false;
        public string? Check;
    }

    // 이름 변경 입력 DTO
    internal sealed class RenameDef
    {
        public string From = "";
        public string To = "";
    }

    // 타입/제약 변경 입력 DTO
    internal sealed class AlterDef
    {
        public string Name = "";
        public string? ToType;      // integer/real/text/bool/json/date
        public bool? ToNotNull;
        public string? ToCheck;
    }

    // GraphQL 응답 DTO
    internal sealed class PresenceItem
    {
        public string? nickname { get; set; }
        public string? sheet { get; set; }
        public string? cell { get; set; }
        public long? updated_at { get; set; } // ms epoch
    }

    internal sealed class PresenceResp
    {
        public PresenceItem[]? presence { get; set; }
    }

    // 서버 tableColumns 결과
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
