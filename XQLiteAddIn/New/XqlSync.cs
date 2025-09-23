// XqlSync.cs  (ExcelPatchApplier 포함 버전)
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Threading;

using Newtonsoft.Json.Linq;

using GraphQL;
using GraphQL.Client.Http;
using GraphQL.Client.Serializer.Newtonsoft;

using ExcelDna.Integration;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace XQLite.AddIn
{
    internal sealed class XqlSync : IDisposable
    {
        private readonly int _pushIntervalMs;
        private readonly int _pullIntervalMs;

        private readonly Backend _backend;
        private readonly XqlMetaRegistry _meta;
        private readonly ConcurrentQueue<EditCell> _outbox = new();
        private readonly object _flushGate = new();
        private readonly object _pullGate = new();

        private long _maxRowVersion;
        public long MaxRowVersion => Interlocked.Read(ref _maxRowVersion);

        private readonly Timer _pushTimer;
        private readonly Timer _pullTimer;

        private volatile bool _started;
        private volatile bool _disposed;

        private readonly ConcurrentQueue<Conflict> _conflicts = new();
        public bool TryDequeueConflict(out Conflict c) => _conflicts.TryDequeue(out c);

        // ⬇️ 새로 추가: 엑셀 반영기
        private readonly ExcelPatchApplier _applier;

        public XqlSync(XqlMetaRegistry meta, string endpoint, string apiKey, int pushIntervalMs = 2000, int pullIntervalMs = 10000)
        {
            _meta = meta ?? throw new ArgumentNullException(nameof(meta));
            _pushIntervalMs = Math.Max(250, pushIntervalMs);
            _pullIntervalMs = Math.Max(1000, pullIntervalMs);

            _backend = new Backend(endpoint, apiKey);
            _applier = new ExcelPatchApplier(_meta);

            _pushTimer = new Timer(_ => SafeFlushUpserts(), null, Timeout.Infinite, Timeout.Infinite);
            _pullTimer = new Timer(_ => SafePull(), null, Timeout.Infinite, Timeout.Infinite);
        }

        public void Start()
        {
            if (_disposed || _started) return;
            _started = true;
            _pushTimer.Change(_pushIntervalMs, _pushIntervalMs);
            _pullTimer.Change(_pullIntervalMs, _pullIntervalMs);

            _backend.StartSubscription(OnServerEvent, MaxRowVersion);
        }

        public void Stop()
        {
            if (!_started) return;
            _started = false;
            _pushTimer.Change(Timeout.Infinite, Timeout.Infinite);
            _pullTimer.Change(Timeout.Infinite, Timeout.Infinite);

            _backend.StopSubscription();
        }

        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;
            try { Stop(); } catch { }
            try { _pushTimer.Dispose(); } catch { }
            try { _pullTimer.Dispose(); } catch { }
            try { _backend.Dispose(); } catch { }
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        public void EnqueueCellEdit(string table, object rowKey, string column, object? value)
        {
            if (_disposed) return;
            _outbox.Enqueue(new EditCell(table, rowKey, column, value));
        }

        public void FlushUpsertsNow() => SafeFlushUpserts();
        public PullResult PullSince(long sinceVersion) => SafePull(sinceVersion);

        private void SafeFlushUpserts()
        {
            if (!_started || _disposed) return;
            lock (_flushGate)
            {
                try { FlushUpsertsCore(); }
                catch (Exception ex) { _conflicts.Enqueue(Conflict.System("flush", ex.Message)); }
            }
        }

        private PullResult SafePull(long? sinceOverride = null)
        {
            if (!_started || _disposed) return PullResult.Empty();
            lock (_pullGate)
            {
                try { return PullCore(sinceOverride ?? MaxRowVersion); }
                catch (Exception ex) { _conflicts.Enqueue(Conflict.System("pull", ex.Message)); return PullResult.Empty(); }
            }
        }

        private void FlushUpsertsCore()
        {
            if (_outbox.IsEmpty) return;
            var batch = DrainDedupCells(_outbox, 512);
            if (batch.Count == 0) return;

            var resp = _backend.UpsertCells(batch);
            if (resp.Errors?.Count > 0)
                foreach (var e in resp.Errors) _conflicts.Enqueue(Conflict.System("upsert", e));

            if (resp.MaxRowVersion > 0)
                InterlockedMax(ref _maxRowVersion, resp.MaxRowVersion);

            if (resp.Conflicts is { Count: > 0 })
                foreach (var c in resp.Conflicts) _conflicts.Enqueue(c);
        }

        private PullResult PullCore(long sinceVersion)
        {
            var resp = _backend.PullRows(sinceVersion);
            if (resp.MaxRowVersion > 0)
                InterlockedMax(ref _maxRowVersion, resp.MaxRowVersion);

            // ⬇️ 서버 패치를 엑셀에 적용 (UI 스레드 매크로 큐로 안전하게)
            if (resp.Patches is { Count: > 0 })
                _applier.ApplyOnUiThread(resp.Patches);

            return resp;
        }

        private void OnServerEvent(ServerEvent ev)
        {
            try
            {
                if (ev.MaxRowVersion > 0)
                    InterlockedMax(ref _maxRowVersion, ev.MaxRowVersion);

                // ⬇️ 푸시 패치 즉시 적용 (UI 스레드)
                if (ev.Patches is { Count: > 0 })
                    _applier.ApplyOnUiThread(ev.Patches);

                // 안전성 위해 한 번 더 Pull
                SafePull();
            }
            catch (Exception ex)
            {
                _conflicts.Enqueue(Conflict.System("subscription", ex.Message));
            }
        }

        private static List<EditCell> DrainDedupCells(ConcurrentQueue<EditCell> q, int max)
        {
            var temp = new List<EditCell>(Math.Min(max * 2, 4096));
            for (int i = 0; i < max && q.TryDequeue(out var e); i++) temp.Add(e);
            if (temp.Count <= 1) return temp;

            var map = new Dictionary<CellKey, EditCell>(temp.Count);
            foreach (var e in temp) map[new CellKey(e.Table, e.RowKey, e.Column)] = e;
            return map.Values.ToList();
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private static void InterlockedMax(ref long target, long value)
        {
            while (true)
            {
                long cur = Volatile.Read(ref target);
                if (value <= cur) return;
                if (Interlocked.CompareExchange(ref target, value, cur) == cur) return;
            }
        }

        private readonly record struct CellKey(string Table, object RowKey, string Column);

        internal readonly record struct EditCell(string Table, object RowKey, string Column, object? Value);

        internal sealed class PullResult
        {
            public long MaxRowVersion { get; set; }
            public List<RowPatch>? Patches { get; set; }
            public static PullResult Empty() => new PullResult { MaxRowVersion = 0, Patches = new List<RowPatch>(0) };
        }

        internal sealed class RowPatch
        {
            public string Table { get; set; } = "";
            public object RowKey { get; set; } = default!;
            public Dictionary<string, object?> Cells { get; set; } = new(StringComparer.Ordinal);
            public long RowVersion { get; set; }
            public bool Deleted { get; set; }
        }

        internal sealed class UpsertResponse
        {
            public long MaxRowVersion { get; set; }
            public List<string>? Errors { get; set; }
            public List<Conflict>? Conflicts { get; set; }
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

        internal sealed class ServerEvent
        {
            public long MaxRowVersion { get; set; }
            public List<RowPatch>? Patches { get; set; }
        }

        // ===== Backend (동일) =====
        private sealed class Backend : IDisposable
        {
            private const string MUT_UPSERT =
@"
mutation($cells:[CellEditInput!]!){
  upsertCells(cells:$cells){
    max_row_version
    errors
    conflicts { table row_key column message server_version local_version }
  }
}";
            private const string Q_PULL =
@"
query($since:Long!){
  rows(since_version:$since){
    max_row_version
    patches { table row_key row_version deleted cells }
  }
}";
            private const string SUB_ROWS =
@"
subscription($since:Long){
  rowsChanged(since_version:$since){
    max_row_version
    patches { table row_key row_version deleted cells }
  }
}";

            private readonly GraphQLHttpClient _http;
            private readonly GraphQLHttpClient _ws;
            private IDisposable? _subscription;

            public Backend(string endpoint, string apiKey)
            {
                if (string.IsNullOrWhiteSpace(endpoint))
                    throw new ArgumentNullException(nameof(endpoint));

                var httpUri = new Uri(endpoint);
                var wsUri = GuessWsUri(httpUri);

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

            public UpsertResponse UpsertCells(IReadOnlyList<EditCell> cells)
            {
                try
                {
                    var vars = new
                    {
                        cells = cells.Select(c => new
                        {
                            table = c.Table,
                            row_key = c.RowKey,
                            column = c.Column,
                            value = c.Value
                        }).ToArray()
                    };

                    var req = new GraphQLRequest { Query = MUT_UPSERT, Variables = vars };
                    var resp = _http.SendQueryAsync<JObject>(req).GetAwaiter().GetResult();
                    return ParseUpsert(resp.Data);
                }
                catch (Exception ex)
                {
                    return new UpsertResponse { MaxRowVersion = 0, Errors = new List<string> { ex.Message } };
                }
            }

            public PullResult PullRows(long since)
            {
                try
                {
                    var req = new GraphQLRequest { Query = Q_PULL, Variables = new { since } };
                    var resp = _http.SendQueryAsync<JObject>(req).GetAwaiter().GetResult();
                    return ParsePull(resp.Data);
                }
                catch (Exception ex)
                {
                    return new PullResult
                    {
                        MaxRowVersion = 0,
                        Patches = new List<RowPatch> {
                            new RowPatch{ Table = "system", RowKey = 0, Cells = new(){ ["error"] = ex.Message }, RowVersion = 0, Deleted = false }
                        }
                    };
                }
            }

            public void StartSubscription(Action<ServerEvent> onEvent, long since)
            {
                StopSubscription();
                try
                {
                    var req = new GraphQLRequest { Query = SUB_ROWS, Variables = new { since } };
                    var observable = _ws.CreateSubscriptionStream<JObject>(req);

                    _subscription = observable.Subscribe(
                        onNext: payload =>
                        {
                            try { onEvent?.Invoke(ParseSub(payload.Data)); }
                            catch { /* ignore */ }
                        },
                        onError: _ =>
                        {
                            try { StopSubscription(); new Timer(_ => StartSubscription(onEvent, since), null, 2000, Timeout.Infinite); }
                            catch { }
                        },
                        onCompleted: () =>
                        {
                            new Timer(_ => StartSubscription(onEvent, since), null, 2000, Timeout.Infinite);
                        });
                }
                catch { /* ignore */ }
            }

            public void StopSubscription()
            {
                try { _subscription?.Dispose(); } catch { }
                _subscription = null;
            }

            private static UpsertResponse ParseUpsert(JObject? data)
            {
                var res = new UpsertResponse { MaxRowVersion = 0, Errors = new List<string>(), Conflicts = new List<Conflict>() };
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
                            Deleted = p["deleted"]?.Type == JTokenType.Boolean && (bool)p["deleted"],
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
                            Deleted = p["deleted"]?.Type == JTokenType.Boolean && (bool)p["deleted"],
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

            private static Uri GuessWsUri(Uri http)
            {
                var scheme = http.Scheme.Equals("https", StringComparison.OrdinalIgnoreCase) ? "wss" : "ws";
                var builder = new UriBuilder(http) { Scheme = scheme };
                return builder.Uri;
            }
        }

        // ========== ⬇️ 엑셀 반영기: 서버 패치 → 시트 적용 (UI 스레드에서 실행) ==========

        private sealed class ExcelPatchApplier
        {
            private readonly XqlMetaRegistry _meta;
            public ExcelPatchApplier(XqlMetaRegistry meta) => _meta = meta;

            public void ApplyOnUiThread(List<RowPatch> patches)
            {
                if (patches == null || patches.Count == 0) return;
                ExcelAsyncUtil.QueueAsMacro(() => { try { ApplyNow(patches); } catch { } });
            }

            private void ApplyNow(List<RowPatch> patches)
            {
                var app = (Excel.Application)ExcelDnaUtil.Application;

                // 테이블별 그룹
                foreach (var grp in patches.GroupBy(p => p.Table, StringComparer.Ordinal))
                {
                    Excel.Worksheet? ws = null;
                    try
                    {
                        ws = FindWorksheetByTable(app, grp.Key, out var smeta);
                        if (ws == null || smeta == null) continue;

                        // 헤더/컬럼 맵
                        var (header, headers) = GetHeaderAndNames(ws);
                        if (headers.Count == 0) continue;

                        // 키 컬럼 인덱스 결정 (메타 우선)
                        int keyCol = FindKeyColumnByMeta(headers, smeta.KeyColumn);
                        int firstDataRow = header.Row + 1;

                        foreach (var patch in grp)
                        {
                            try
                            {
                                int? row = FindRowByKey(ws, firstDataRow, keyCol, patch.RowKey);
                                if (patch.Deleted)
                                {
                                    if (row.HasValue) SafeDeleteRow(ws, row.Value);
                                    continue;
                                }
                                if (!row.HasValue) row = AppendNewRow(ws, firstDataRow);

                                ApplyCells(ws, row!.Value, headers, smeta, patch.Cells);
                            }
                            catch { /* per-row safe */ }
                        }
                    }
                    finally { ReleaseCom(ws); }
                }
            }

            // === 메타 기반: 테이블명 → 워크시트 찾기 ===
            private Excel.Worksheet? FindWorksheetByTable(Excel.Application app, string table, out SheetMeta? smeta)
            {
                smeta = null;

                // 1) 시트명 == 테이블명인 경우
                try
                {
                    foreach (Excel.Worksheet w in app.Worksheets)
                    {
                        try
                        {
                            string name = w.Name;
                            // 메타가 등록된 시트만 대상
                            if (_meta.TryGetSheet(name, out var m))
                            {
                                if (string.Equals(m.TableName ?? name, table, StringComparison.Ordinal))
                                {
                                    smeta = m;
                                    return w;
                                }
                                // 시트명 자체가 테이블명인 케이스도 통과
                                if (string.Equals(name, table, StringComparison.Ordinal) && smeta == null)
                                {
                                    smeta = m;
                                    return w;
                                }
                            }
                        }
                        finally { ReleaseCom(w); }
                    }
                }
                catch { }

                return null;
            }

            private static (Excel.Range header, List<string> headers) GetHeaderAndNames(Excel.Worksheet ws)
            {
                Excel.Range? header = null;
                var names = new List<string>();
                try
                {
                    header = ws.Range[ws.Cells[1, 1], ws.Cells[1, ws.UsedRange.Columns.Count]];
                    int cols = header.Columns.Count;
                    for (int c = 1; c <= cols; c++)
                    {
                        string name = "";
                        try { name = Convert.ToString(((Excel.Range)header.Cells[1, c]).Value2) ?? ""; }
                        catch { }
                        names.Add(name.Trim());
                    }
                    return (header, names);
                }
                catch { return ((Excel.Range)ws.Cells[1, 1], names); }
                finally { ReleaseCom(header); }
            }

            private static int FindKeyColumnByMeta(List<string> headers, string keyName)
            {
                if (!string.IsNullOrWhiteSpace(keyName))
                {
                    var idx = headers.FindIndex(h => string.Equals(h, keyName, StringComparison.Ordinal));
                    if (idx >= 0) return idx + 1; // 1-based
                }
                // fallback: id/key/첫번째
                var id = headers.FindIndex(h => string.Equals(h, "id", StringComparison.OrdinalIgnoreCase));
                if (id >= 0) return id + 1;
                var key = headers.FindIndex(h => string.Equals(h, "key", StringComparison.OrdinalIgnoreCase));
                if (key >= 0) return key + 1;
                return 1;
            }

            private static int? FindRowByKey(Excel.Worksheet ws, int firstDataRow, int keyCol, object key)
            {
                try
                {
                    var used = ws.UsedRange;
                    int lastRow = used.Row + used.Rows.Count - 1;
                    ReleaseCom(used);

                    for (int r = firstDataRow; r <= lastRow; r++)
                    {
                        Excel.Range? cell = null;
                        try
                        {
                            cell = (Excel.Range)ws.Cells[r, keyCol];
                            var v = cell.Value2;
                            if (EqualKey(v, key)) return r;
                        }
                        catch { }
                        finally { ReleaseCom(cell); }
                    }
                }
                catch { }
                return null;
            }

            private static bool EqualKey(object? excelVal, object key)
            {
                if (excelVal == null) return key == null;
                if (excelVal is double d)
                {
                    if (key is double kd) return Math.Abs(d - kd) < 1e-9;
                    if (key is long kl) return Math.Abs(d - kl) < 1e-9;
                    if (key is int ki) return Math.Abs(d - ki) < 1e-9;
                    if (double.TryParse(Convert.ToString(key), out var kdp)) return Math.Abs(d - kdp) < 1e-9;
                }
                var s1 = Convert.ToString(excelVal)?.Trim();
                var s2 = Convert.ToString(key)?.Trim();
                return string.Equals(s1, s2, StringComparison.Ordinal);
            }

            private static int AppendNewRow(Excel.Worksheet ws, int firstDataRow)
            {
                int last = firstDataRow;
                try
                {
                    var used = ws.UsedRange;
                    last = used.Row + used.Rows.Count - 1;
                    ReleaseCom(used);
                }
                catch { }
                return Math.Max(firstDataRow, last + 1);
            }

            // ✅ 메타 컬럼에 정의된 컬럼만 적용
            private static void ApplyCells(Excel.Worksheet ws, int row, List<string> headers, SheetMeta meta, Dictionary<string, object?> cells)
            {
                for (int c = 0; c < headers.Count; c++)
                {
                    var colName = headers[c];
                    if (string.IsNullOrWhiteSpace(colName)) continue;
                    if (!meta.Columns.ContainsKey(colName)) continue; // 메타에 없는 컬럼은 skip

                    if (!cells.TryGetValue(colName, out var val)) continue;

                    Excel.Range? rg = null;
                    try
                    {
                        rg = (Excel.Range)ws.Cells[row, c + 1];
                        if (val == null) { rg.Value2 = null; continue; }

                        switch (val)
                        {
                            case bool b: rg.Value2 = b; break;
                            case long l: rg.Value2 = (double)l; break;
                            case int i: rg.Value2 = (double)i; break;
                            case double d: rg.Value2 = d; break;
                            case float f: rg.Value2 = (double)f; break;
                            case decimal m: rg.Value2 = (double)m; break;
                            case DateTime dt: rg.Value2 = dt; break;
                            default: rg.Value2 = val.ToString(); break;
                        }
                    }
                    catch { }
                    finally { ReleaseCom(rg); }
                }
            }

            private static void SafeDeleteRow(Excel.Worksheet ws, int row)
            {
                try { var rg = (Excel.Range)ws.Rows[row]; rg.Delete(); ReleaseCom(rg); }
                catch { }
            }

            private static void ReleaseCom(object? o)
            {
                try { if (o != null && Marshal.IsComObject(o)) Marshal.FinalReleaseComObject(o); }
                catch { }
            }
        }
    }
}
