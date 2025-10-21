// XqlGqlBackend.cs (async-first, index.ts 프로토콜 정렬 + 안정화)

using GraphQL;
using GraphQL.Client.Http;
using GraphQL.Client.Serializer.Newtonsoft;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
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

        // Sync
        Task<UpsertResult> UpsertCells(IEnumerable<EditCell> cells, CancellationToken ct = default);
        Task<PullResult> PullRows(long since, CancellationToken ct = default);
        void StartSubscription(Action<ServerEvent> onEvent, long since);
        void StopSubscription();

        // Collab
        Task PresenceTouch(string nickname, string? sheet, string? cell, CancellationToken ct = default);
        Task AcquireLock(string cellOrResourceKey, string by, CancellationToken ct = default);
        Task ReleaseLocksBy(string by, CancellationToken ct = default);

        // Schema
        Task TryCreateTable(string table, string key, CancellationToken ct = default);
        Task TryAddColumns(string table, IEnumerable<ColumnDef> cols, CancellationToken ct = default);
        Task TryDropColumns(string table, IEnumerable<string> names, CancellationToken ct = default);
        Task TryRenameColumns(string table, IEnumerable<RenameDef> renames, CancellationToken ct = default);
        Task TryAlterColumns(string table, IEnumerable<AlterDef> alters, CancellationToken ct = default);

        // Meta / Audit / Export
        Task<JObject?> TryFetchServerMeta(CancellationToken ct = default);
        Task<JArray?> TryFetchAuditLog(long? since = null, CancellationToken ct = default);
        Task<byte[]?> TryExportDatabase(CancellationToken ct = default);

        // Presence list
        Task<PresenceItem[]?> FetchPresence(CancellationToken ct = default);

        // Row batch upsert (행 JSON) — 서버가 PK 자동발급 가능
        Task<UpsertResult> UpsertRows(string table, List<Dictionary<string, object?>> rows, CancellationToken ct = default);

        // 상태 하트비트
        Task<long> Ping(CancellationToken ct = default);

        // Table columns
        Task<List<ColumnInfo>> GetTableColumns(string table, CancellationToken ct = default);

        // Rows snapshot(JSON 배열 → RowPatch 변환)
        Task<List<RowPatch>> FetchRowsSnapshot(string table, CancellationToken ct = default);

        // 상태
        event Action<ConnState, string?>? StateChanged;
        ConnState State { get; }
        string? StateDetail { get; }
        DateTime LastOkUtc { get; }
    }

    // =======================================================================
    // GraphQL 구현
    // =======================================================================
    internal sealed class XqlGqlBackend : IXqlBackend
    {
        // ── GraphQL 문서 (index.ts 스키마 기준) ─────────────────────────────

        private const string Q_PULL = @"
        query($since:Int!){
          rows(since_version:$since){
            max_row_version
            patches { table row_key row_version deleted cells }
          }
        }";

        private const string SUB_ROWS = @"
        subscription{
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
            assigned { table temp_row_key new_id }
          }
        }";

        private const string MUT_PRESENCE = @"
        mutation($n:String!,$s:String,$c:String){
          presenceTouch(nickname:$n, sheet:$s, cell:$c){ ok }
        }";

        private const string MUT_ACQUIRE = @"mutation($cell:String!, $by:String!){ acquireLock(cell:$cell, by:$by){ ok } }";
        private const string MUT_RELEASE_BY = @"mutation($by:String!){ releaseLocksBy(by:$by){ ok } }";

        private const string MUT_CREATE_TABLE = @"
        mutation($table:String!, $key:String!){
          createTable(table:$table, key:$key){ ok }
        }";

        private const string MUT_ADD_COLUMNS = @"
        mutation($table:String!, $columns:[ColumnDefInput!]!){
          addColumns(table:$table, columns:$columns){ ok }
        }";

        private const string MUT_DROP_COLUMNS = @"
        mutation($t:String!,$ns:[String!]!){
          dropColumns(table:$t, names:$ns){ ok }
        }";

        private const string MUT_RENAME_COLUMNS = @"
        mutation($t:String!,$rs:[RenameDefInput!]!){
          renameColumns(table:$t, renames:$rs){ ok }
        }";

        private const string MUT_ALTER_COLUMNS = @"
        mutation($t:String!,$alters:[AlterDefInput!]!){
          alterColumns(table:$t, alters:$alters){ ok }
        }";

        private const string MUT_UPSERT_ROWS = @"
        mutation($table:String!,$rows:[JSON!]!){
          upsertRows(table:$table, rows:$rows){
            max_row_version
            errors
            assigned { table temp_row_key new_id }
          }
        }";

        private const string Q_META = @"query{ meta }";

        private const string Q_AUDIT = @"
        query($since:Int){
          audit_log(since_version:$since){
            ts user table row_key column old_value new_value row_version
          }
        }";

        private const string Q_EXPORT_DB = @"query{ exportDatabase }";
        private const string Q_PRESENCE = @"query { presence { nickname sheet cell updated_at } }";

        private const string Q_TABLE_COLUMNS = @"
        query($t:String!){
          tableColumns(table:$t){ name type notnull pk }
        }";

        private const string Q_ROWS_SNAPSHOT = @"
        query($t:String!){
          rowsSnapshot(table:$t)
        }";

        // ── 필드 ─────────────────────────────────────────────────────────────
        private readonly GraphQLHttpClient _http;
        private readonly GraphQLHttpClient _ws;
        private IDisposable? _subscription;
        private int _subRetry;

        private readonly Timer _heartbeat;
        private Timer? _resubTimer;

        public IXqlBackend.ConnState State { get; private set; } = IXqlBackend.ConnState.Connecting;
        public string? StateDetail { get; private set; }
        public DateTime LastOkUtc { get; private set; } = DateTime.MinValue;
        public event Action<IXqlBackend.ConnState, string?>? StateChanged;

        private const int HB_TTL_MS = 10_000;

        // ── 생성자 ───────────────────────────────────────────────────────────
        internal XqlGqlBackend(string httpEndpoint, string? apiKey, string? project = null, int heartbeatSec = 3)
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

            // project 헤더
            var prj = string.IsNullOrWhiteSpace(project) ? "default" : project!.Trim();
            const string H = "x-project";
            try
            {
                if (_http.HttpClient.DefaultRequestHeaders.Contains(H)) _http.HttpClient.DefaultRequestHeaders.Remove(H);
                if (_ws.HttpClient.DefaultRequestHeaders.Contains(H)) _ws.HttpClient.DefaultRequestHeaders.Remove(H);
                _http.HttpClient.DefaultRequestHeaders.Add(H, prj);
                _ws.HttpClient.DefaultRequestHeaders.Add(H, prj);
            }
            catch { /* ignore */ }

            _heartbeat = new Timer(async _ => await SafeHeartbeat().ConfigureAwait(false), null, Timeout.Infinite, Timeout.Infinite);
            _ = SafeHeartbeat();
            _heartbeat.Change(TimeSpan.FromSeconds(heartbeatSec), TimeSpan.FromSeconds(heartbeatSec));
        }

        public void Dispose()
        {
            try { StopSubscription(); } catch { }
            try { _resubTimer?.Dispose(); _resubTimer = null; } catch { }
            try { _http.Dispose(); } catch { }
            try { _ws.Dispose(); } catch { }
            try { _heartbeat.Change(Timeout.Infinite, Timeout.Infinite); _heartbeat.Dispose(); } catch { }
        }

        // ── 상태 갱신 ────────────────────────────────────────────────────────
        private void SetState(IXqlBackend.ConnState st, string? detail = null)
        {
            if (State == st && detail == StateDetail) return;
            State = st;
            StateDetail = detail;
            if (st == IXqlBackend.ConnState.Online) LastOkUtc = DateTime.UtcNow;
            try { StateChanged?.Invoke(st, detail); } catch { }
        }

        public async Task<long> Ping(CancellationToken ct = default)
        {
            try
            {
                var req = new GraphQLHttpRequest { Query = "query { ping }" };
                var resp = await _http.SendQueryAsync<PingQueryResult>(req, ct).ConfigureAwait(false);
                var now = resp.Data?.ping ?? 0L;
                SetState(IXqlBackend.ConnState.Online, "ping ok");
                return now;
            }
            catch (Exception ex)
            {
                var now = DateTime.UtcNow;
                if (LastOkUtc == DateTime.MinValue)
                {
                    SetState(IXqlBackend.ConnState.Connecting, "ping...");
                }
                else
                {
                    var ms = (now - LastOkUtc).TotalMilliseconds;
                    if (ms > HB_TTL_MS) SetState(IXqlBackend.ConnState.Disconnected, $"ping timeout {(int)(ms / 1000)}s");
                    else SetState(IXqlBackend.ConnState.Degraded, $"ping fail: {ex.GetType().Name}");
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
            var since32 = (int)XqlCommon.Clamp(since, int.MinValue, int.MaxValue);
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
                p => { _subRetry = 0; try { onEvent(ParseSub(p.Data)); } catch { } },
                _ => Resubscribe(onEvent, since),
                () => Resubscribe(onEvent, since)
            );
            Interlocked.Exchange(ref _subscription, sub)?.Dispose();
        }

        private void Resubscribe(Action<ServerEvent> onEvent, long since)
        {
            StopSubscription();
            var delayMs = (int)Math.Min(30_000, 500 * Math.Pow(2, Math.Min(_subRetry++, 10)));
            _resubTimer?.Dispose();
            _resubTimer = new Timer(_ =>
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

        public Task AcquireLock(string cellOrResourceKey, string by, CancellationToken ct = default)
            => _http.SendMutationAsync<JObject>(new GraphQLRequest { Query = MUT_ACQUIRE, Variables = new { cell = cellOrResourceKey, by } }, ct);

        public Task ReleaseLocksBy(string by, CancellationToken ct = default)
            => _http.SendMutationAsync<JObject>(new GraphQLRequest { Query = MUT_RELEASE_BY, Variables = new { by } }, ct);

        // ── Schema ───────────────────────────────────────────────────────────
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
                                      type = MapType(c.Kind),
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

        // ── Meta / Audit / Export / Presence ────────────────────────────────
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

        public async Task<PresenceItem[]?> FetchPresence(CancellationToken ct = default)
        {
            var req = new GraphQLRequest { Query = Q_PRESENCE };
            var resp = await _http.SendQueryAsync<PresenceResp>(req, ct).ConfigureAwait(false);
            var arr = resp.Data?.presence ?? Array.Empty<PresenceItem>();
            return arr.Where(p => p != null).ToArray();
        }

        // 행 단위 업서트 (PK 미포함 → 서버 자동발급 가정, 반영은 Pull로 수신)
        public async Task<UpsertResult> UpsertRows(string table, List<Dictionary<string, object?>> rows, CancellationToken ct = default)
        {
            var req = new GraphQLRequest { Query = MUT_UPSERT_ROWS, Variables = new { table, rows } };
            var resp = await _http.SendMutationAsync<JObject>(req, ct).ConfigureAwait(false);
            return ParseUpsert(resp.Data);
        }

        // 스냅샷(JSON 배열) → RowPatch 변환
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
                    var deleted = it["deleted"]?.Type == JTokenType.Integer
                                  ? ((int?)it["deleted"] ?? 0) != 0
                                  : (it["deleted"]?.Type == JTokenType.Boolean ? ((bool?)it["deleted"] ?? false) : false);

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
                    list.Add(new RowPatch { Table = table, RowKey = rk, Deleted = deleted, Cells = cells });
                }
                return list;
            }
            catch
            {
                return new List<RowPatch>();
            }
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
                Conflicts = root["conflicts"] is JArray ca ? ca.ToObject<List<Conflict>>() : null,
                Assigned = root["assigned"] is JArray aa ? aa
                    .OfType<JObject>()
                    .Select(a => new AssignedId
                    {
                        Table = a["table"]?.ToString() ?? "",
                        TempRowKey = a["temp_row_key"]?.ToString(),
                        NewId = a["new_id"]?.ToString() ?? ""
                    })
                    .Where(x => !string.IsNullOrWhiteSpace(x.NewId))
                    .ToList() : null
            };
        }

        public async Task<List<ColumnInfo>> GetTableColumns(string table, CancellationToken ct = default)
        {
            var req = new GraphQLRequest { Query = Q_TABLE_COLUMNS, Variables = new { t = table } };
            var resp = await _http.SendQueryAsync<JObject>(req, ct).ConfigureAwait(false);
            var arr = resp.Data?["tableColumns"] as JArray;
            return arr?.ToObject<List<ColumnInfo>>() ?? new List<ColumnInfo>();
        }

        // ── Parser ───────────────────────────────────────────────────────────
        private static UpsertResult ParseUpsert(JObject? data)
        {
            var res = new UpsertResult { MaxRowVersion = 0, Errors = new List<string>(), Conflicts = new List<Conflict>(), Assigned = null };
            if (data == null) return res;

            var root = data["upsertCells"] ?? data["upsertRows"];
            if (root is not JObject u) return res;

            res.MaxRowVersion = (long?)u["max_row_version"] ?? 0;

            if (u["assigned"] is JArray aa)
            {
                res.Assigned = aa
                    .OfType<JObject>()
                    .Select(a => new AssignedId
                    {
                        Table = a["table"]?.ToString() ?? "",
                        TempRowKey = a["temp_row_key"]?.ToString(),
                        NewId = a["new_id"]?.ToString() ?? ""
                    })
                    .Where(x => !string.IsNullOrWhiteSpace(x.NewId))
                    .ToList();
            }

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

        private static bool ParseDeleted(JToken? tok)
        {
            if (tok == null || tok.Type == JTokenType.Null) return false;
            if (tok.Type == JTokenType.Boolean) return (bool)tok!;
            if (tok.Type == JTokenType.Integer) return ((long)tok!) != 0;
            if (tok.Type == JTokenType.String && long.TryParse((string)tok!, out var n)) return n != 0;
            return false;
        }

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
                case "date": return "integer"; // epoch ms
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

    internal sealed class ColumnDef
    {
        public string Name = "";
        public string Kind = "text";
        public bool NotNull = false;
        public string? Check;
    }

    internal sealed class RenameDef
    {
        public string From = "";
        public string To = "";
    }

    internal sealed class AlterDef
    {
        public string Name = "";
        public string? ToType;
        public bool? ToNotNull;
        public string? ToCheck;
    }

    internal sealed class PresenceItem
    {
        public string? nickname { get; set; }
        public string? sheet { get; set; }
        public string? cell { get; set; }
        public long? updated_at { get; set; }
    }

    internal sealed class PresenceResp
    {
        public PresenceItem[]? presence { get; set; }
    }

    public sealed class ColumnInfo
    {
        public string name { get; set; } = "";
        public string? type { get; set; }
        public bool notnull { get; set; }
        public bool pk { get; set; }
    }

    internal sealed class AssignedId
    {
        public string Table { get; set; } = "";
        public string? TempRowKey { get; set; }
        public string NewId { get; set; } = "";
    }

    internal sealed class UpsertResult
    {
        public long MaxRowVersion;
        public List<string>? Errors;
        public List<Conflict>? Conflicts;
        public List<AssignedId>? Assigned;
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
