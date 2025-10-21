// XqlSync.cs  (ExcelPatchApplier 포함 버전)
using ExcelDna.Integration;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace XQLite.AddIn
{
    internal sealed class XqlSync : IDisposable
    {
        private sealed class PersistentState
        {
            public string LastSessionId { get; set; } = "";
            public string Project { get; set; } = "";
            public string Workbook { get; set; } = "";
            public long LastMaxRowVersion { get; set; } = 0;
            public DateTime LastFullPullUtc { get; set; } = DateTime.MinValue;
            public string? LastSchemaHash { get; set; }
            public DateTime LastMetaUtc { get; set; } = DateTime.MinValue;
        }

        private readonly int _pushIntervalMs;
        private readonly int _pullIntervalMs;

        private readonly IXqlBackend _backend;
        private readonly XqlSheet _sheet;
        private readonly ConcurrentQueue<EditCell> _outbox = new();
        private readonly SemaphoreSlim _pushSem = new(1, 1);
        private readonly SemaphoreSlim _pullSem = new(1, 1);
        private int _pulling; // 0/1 (Interlocked)
        public bool IsPulling => System.Threading.Volatile.Read(ref _pulling) == 1;
        public event Action<bool>? PullStateChanged; // true=시작, false=종료
        private long _pullBackoffUntilMs;
        private int _pullErr; // 연속 오류 횟수

        private long _maxRowVersion;
        public long MaxRowVersion => Interlocked.Read(ref _maxRowVersion);

        private readonly Timer _pushTimer;
        private readonly Timer _pullTimer;

        private volatile bool _started;
        private volatile bool _disposed;

        private const int UPSERT_CHUNK = 512;    // 1회 전송 셀 수
        private const int UPSERT_SLICE_MS = 250; // 한번에 잡는 시간
        private const int LAST_PUSHED_MAX = 100_000;
        private readonly LinkedList<string> _lruKeys = new();
        private readonly Dictionary<string, (string? val, LinkedListNode<string> node)> _lastPushedLru
            = new(StringComparer.Ordinal);

        private const int CONFLICT_MAX = 5000;

        private readonly ConcurrentQueue<Conflict> _conflicts = new();

        private string? _workbookFullName;
        private PersistentState _state = new();
        private volatile bool _forceFullPull = false;

        private CancellationTokenSource _cts = new();

        public XqlSync(IXqlBackend backend, XqlSheet sheet, int pushIntervalMs = 2000, int pullIntervalMs = 10000)
        {
            _sheet = sheet ?? throw new ArgumentNullException(nameof(sheet));
            _pushIntervalMs = Math.Max(250, pushIntervalMs);
            _pullIntervalMs = Math.Max(1000, pullIntervalMs);

            _backend = backend ?? throw new ArgumentNullException(nameof(backend));

            _pushTimer = new Timer(_ => SafeFlushUpserts(), null, Timeout.Infinite, Timeout.Infinite);
            _pullTimer = new Timer(_ => _ = SafePull(), null, Timeout.Infinite, Timeout.Infinite);
        }

        public void Start()
        {
            if (_disposed || _started) return;
            _started = true;

            _cts = new CancellationTokenSource();

            _pushTimer.Change(_pushIntervalMs, _pushIntervalMs);
            _pullTimer.Change(_pullIntervalMs, _pullIntervalMs);

            // 서버 이벤트 구독(동기 진입, 내부에서 별도 스레드)
            _backend.StartSubscription(OnServerEvent, MaxRowVersion);
        }

        public void Stop()
        {
            if (!_started) return;
            _started = false;

            try { _cts.Cancel(); } catch { }

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
        }

        private static string Key(EditCell e) => $"{e.Table}\n{XqlCommon.ValueToString(e.RowKey)}\n{e.Column}";

        public void EnqueueIfChanged(string table, string rowKey, string column, object? value)
        {
            var e = new EditCell(Table: table, RowKey: rowKey, Column: column, Value: value);
            var k = Key(e);
            var norm = XqlCommon.Canonicalize(value);
            if (IsSameAsLast(k, norm)) return;
            _outbox.Enqueue(e);
        }

        public bool TryDequeueConflict(out Conflict c) => _conflicts.TryDequeue(out c);

        // 초기화 진입점 (워크북 오픈 시 1회)
        public void InitPersistentState(string workbookFullName, string? project = null)
        {
            _workbookFullName = workbookFullName;
            var wbName = Path.GetFileNameWithoutExtension(workbookFullName) ?? "wb";

            var proj = (project ?? XqlConfig.Project ?? "").Trim();
            if (string.IsNullOrEmpty(proj)) proj = wbName; // 비면 워크북명으로

            _state = new PersistentState
            {
                Project = proj,
                Workbook = wbName,
                LastMaxRowVersion = 0,
                LastFullPullUtc = DateTime.MinValue
            };

            // 워크북에서 K/V 읽기 (UI 스레드에서 안전하게)
            Dictionary<string, string> loaded = new(StringComparer.Ordinal);
            try
            {
                loaded = XqlCommon.OnExcelThreadAsync(() =>
                {
                    var app = (Excel.Application)ExcelDnaUtil.Application;
                    Excel.Workbook? wb = null;
                    try
                    {
                        foreach (Excel.Workbook w in app.Workbooks)
                        {
                            try
                            {
                                if (string.Equals(w.FullName, workbookFullName, StringComparison.OrdinalIgnoreCase))
                                { wb = w; break; }
                            }
                            finally { if (!ReferenceEquals(wb, w)) XqlCommon.ReleaseCom(w); }
                        }
                        wb ??= app.ActiveWorkbook;
                        return wb != null ? XqlSheet.StateReadAll(wb) : new Dictionary<string, string>(StringComparer.Ordinal);
                    }
                    finally { XqlCommon.ReleaseCom(wb); }
                }).GetAwaiter().GetResult();
            }
            catch
            {
                // Excel이 바쁜 경우 등 — 빈 상태로 진행
                loaded = new Dictionary<string, string>(StringComparer.Ordinal);
            }

            // 값 반영
            if (loaded.TryGetValue("last_max_row_version", out var s) && long.TryParse(s, out var l))
                _state.LastMaxRowVersion = l;
            if (loaded.TryGetValue("last_schema_hash", out var h)) _state.LastSchemaHash = h;
            if (loaded.TryGetValue("last_full_pull_utc", out var f) && DateTime.TryParse(f, out var dt))
                _state.LastFullPullUtc = dt;

            // 새 세션 시작
            _forceFullPull = XqlConfig.AlwaysFullPullOnStartup;
            _state.LastSessionId = Guid.NewGuid().ToString("N");
            PersistState();

            // (선택) 서버 메타 확인 → 스키마 변경 감지 시 Full Pull 예약
            if (XqlConfig.FullPullWhenSchemaChanged)
            {
                Task.Run(async () =>
                {
                    try
                    {
                        var meta = await _backend.TryFetchServerMeta().ConfigureAwait(false);
                        var hash = meta?["schema_hash"]?.ToString();
                        if (!string.IsNullOrWhiteSpace(hash) && !string.Equals(hash, _state.LastSchemaHash, StringComparison.Ordinal))
                            _forceFullPull = true;
                        _state.LastSchemaHash = hash;
                        _state.LastMetaUtc = DateTime.UtcNow;
                        PersistState();
                    }
                    catch { }
                });
            }
        }

        public void FlushUpsertsNow() => _ = FlushUpsertsNow(false);

        public async Task PullSince(long? sinceOverride = null)
        {
            if (XqlCommon.NowMs() < _pullBackoffUntilMs) return;
            if (Interlocked.Exchange(ref _pulling, 1) == 1) return;

            PullStateChanged?.Invoke(true);
            if (!await _pullSem.WaitAsync(0).ConfigureAwait(false))
            {
                Interlocked.Exchange(ref _pulling, 0);
                PullStateChanged?.Invoke(false);
                return;
            }

            try
            {
                var since = sinceOverride ?? (_forceFullPull ? 0 : MaxRowVersion);
                var pr = await _backend.PullRows(since, _cts.Token).ConfigureAwait(false);

                // 초기 상태: 패치가 없고 로컬 max==0 이면 헤더 부트스트랩
                if ((pr.Patches == null || pr.Patches.Count == 0) && (since == 0 || MaxRowVersion == 0))
                {
                    await ApplyBootstrapAsync(pr).ConfigureAwait(false);
                    _pullErr = 0;
                    _forceFullPull = false;
                    return;
                }

                // 증분 패치 적용
                await ApplyIncrementalPatches(pr).ConfigureAwait(false);

                if (pr.MaxRowVersion > 0)
                    XqlCommon.InterlockedMax(ref _maxRowVersion, pr.MaxRowVersion);

                _pullErr = 0;
                _pullBackoffUntilMs = 0;
                _state.LastMaxRowVersion = MaxRowVersion;
                PersistState();
            }
            catch
            {
                _pullErr = Math.Min(_pullErr + 1, 4);
                _pullBackoffUntilMs = XqlCommon.NowMs() + _pullErr * 2000L;
            }
            finally
            {
                _pullSem.Release();
                Interlocked.Exchange(ref _pulling, 0);
                PullStateChanged?.Invoke(false);
            }
        }

        public async Task FlushUpsertsNow(bool force = false)
        {
            if (_disposed) return;

            // force면 _started 여부와 무관하게 1회 실행
            if (!force && (!_started)) return;

            if (!_pushSem.Wait(0)) return;
            try { await FlushUpsertsCore().ConfigureAwait(false); }
            finally { _pushSem.Release(); }
        }

        // Private

        // 풀 전체가 처음이거나, 증분이 0건인 초기 상태에서 호출
        private async Task ApplyBootstrapAsync(PullResult pr)
        {
            // 1) 서버 메타 가져와서 '플랜'만 구성 (비-COM)
            var meta = await _backend.TryFetchServerMeta().ConfigureAwait(false);
            if (meta == null) return;

            var schema = new Dictionary<string, List<string>>(StringComparer.Ordinal);

            // (A) 신형 포맷: meta.tables[{ name, key, columns }]
            if (meta["tables"] is JArray tablesNew)
            {
                foreach (var t in tablesNew.OfType<JObject>())
                {
                    var tname = (t["name"] ?? t["table_name"])?.ToString();
                    if (string.IsNullOrWhiteSpace(tname)) continue;

                    // columns: [ {name:..}, ... ] or [ "id","Name", ... ]
                    var cols = new List<string>();
                    if (t["columns"] is JArray ca)
                    {
                        foreach (var e in ca)
                        {
                            if (e is JObject jo && jo["name"] != null)
                                cols.Add((jo["name"]!.ToString() ?? "").Trim());
                            else
                                cols.Add((e?.ToString() ?? "").Trim());
                        }
                    }

                    // 컬럼이 비어있으면 서버에 질의(레거시/보강)
                    if (cols.Count == 0)
                    {
                        try
                        {
                            var infos = await _backend.GetTableColumns(tname!).ConfigureAwait(false);
                            cols = infos.Select(i => (i.name ?? "").Trim())
                                        .Where(s => s.Length > 0)
                                        .Distinct(StringComparer.OrdinalIgnoreCase).ToList();
                        }
                        catch { /* ignore */ }
                    }

                    // 키 보강(id 선호)
                    var key = (t["key"] ?? t["key_column"])?.ToString();
                    key = string.IsNullOrWhiteSpace(key) ? "id" : key!;
                    if (!cols.Any(c => c.Equals(key, StringComparison.OrdinalIgnoreCase)))
                        cols.Insert(0, key);

                    cols = cols.Where(s => !string.IsNullOrWhiteSpace(s))
                               .Distinct(StringComparer.OrdinalIgnoreCase)
                               .ToList();

                    if (cols.Count > 0) schema[tname!] = cols;
                }
            }
            // (B) 레거시 포맷: meta.schema[{ table_name, key_column }]
            else if (meta["schema"] is JArray legacy)
            {
                foreach (var t in legacy.OfType<JObject>())
                {
                    var tname = t["table_name"]?.ToString();
                    if (string.IsNullOrWhiteSpace(tname)) continue;

                    var cols = new List<string>();
                    try
                    {
                        var infos = await _backend.GetTableColumns(tname!).ConfigureAwait(false);
                        cols = infos.Select(i => (i.name ?? "").Trim())
                                    .Where(s => s.Length > 0)
                                    .Distinct(StringComparer.OrdinalIgnoreCase).ToList();
                    }
                    catch { /* ignore */ }

                    var key = t["key_column"]?.ToString();
                    key = string.IsNullOrWhiteSpace(key) ? "id" : key!;
                    if (!cols.Any(c => c.Equals(key, StringComparison.OrdinalIgnoreCase)))
                        cols.Insert(0, key);

                    if (cols.Count > 0) schema[tname!] = cols;
                }
            }

            if (schema.Count == 0) return;

            // 2) UI 스레드에서만 Excel 만지기
            XqlSheetView.ApplyPlanAndPatches(schema, pr?.Patches);
        }

        // 증분 패치를 UI 스레드에서 적용 (항상 한 경로)
        private Task ApplyIncrementalPatches(PullResult pr)
        {
            if (pr?.Patches is { Count: > 0 })
                XqlSheetView.ApplyPlanAndPatches(null, pr.Patches); // 내부에서 UI 스레드 마샬링
            return Task.CompletedTask;
        }

        private void RememberPushed(string k, string? v)
        {
            if (_lastPushedLru.TryGetValue(k, out var ent))
            {
                ent.val = v;
                _lruKeys.Remove(ent.node);
                _lruKeys.AddFirst(ent.node);
                _lastPushedLru[k] = (v, ent.node);
                return;
            }
            var node = new LinkedListNode<string>(k);
            _lruKeys.AddFirst(node);
            _lastPushedLru[k] = (v, node);

            if (_lastPushedLru.Count > LAST_PUSHED_MAX)
            {
                var tail = _lruKeys.Last;
                if (tail != null)
                {
                    _lastPushedLru.Remove(tail.Value);
                    _lruKeys.RemoveLast();
                }
            }
        }

        private bool IsSameAsLast(string k, string? v)
        {
            return _lastPushedLru.TryGetValue(k, out var ent) && ent.val == v;
        }

        private void PersistState()
        {
            try
            {
                if (_workbookFullName == null) return;
                var kv = new Dictionary<string, string>(StringComparer.Ordinal)
                {
                    ["last_session_id"] = _state.LastSessionId ?? "",
                    ["project"] = _state.Project ?? "",
                    ["workbook"] = _state.Workbook ?? "",
                    ["last_max_row_version"] = _state.LastMaxRowVersion.ToString(CultureInfo.InvariantCulture),
                    ["last_full_pull_utc"] = (_state.LastFullPullUtc == DateTime.MinValue ? "" : _state.LastFullPullUtc.ToString("o")),
                    ["last_schema_hash"] = _state.LastSchemaHash ?? "",
                    ["last_meta_utc"] = (_state.LastMetaUtc == DateTime.MinValue ? "" : _state.LastMetaUtc.ToString("o")),
                };

                // UI 스레드에서만 상태 기록
                _ = XqlCommon.OnExcelThreadAsync(() =>
                {
                    var app = (Excel.Application)ExcelDnaUtil.Application;
                    Excel.Workbook? wb = null;
                    try
                    {
                        foreach (Excel.Workbook w in app.Workbooks)
                        {
                            try
                            {
                                if (string.Equals(w.FullName, _workbookFullName, StringComparison.OrdinalIgnoreCase))
                                { wb = w; break; }
                            }
                            finally { if (!ReferenceEquals(wb, w)) XqlCommon.ReleaseCom(w); }
                        }
                        wb ??= app.ActiveWorkbook;
                        if (wb != null)
                            XqlSheet.StateSetMany(wb, kv);
                    }
                    finally { XqlCommon.ReleaseCom(wb); }
                    return 0;
                });
            }
            catch { }
        }

        private void SafeFlushUpserts()
        {
            if (!_started || _disposed) return;

            try
            {
                if (!_pushSem.Wait(0)) return;
                _ = Task.Run(async () =>
                {
                    try { await FlushUpsertsCore(); }
                    catch (Exception ex) { PushConflict(Conflict.System("upsert.core", ex.Message)); }
                    finally { _pushSem.Release(); }
                });
            }
            catch (Exception ex)
            {
                PushConflict(Conflict.System("flush", ex.Message));
            }
        }

        private void PushConflict(Conflict c)
        {
            _conflicts.Enqueue(c);
            while (_conflicts.Count > CONFLICT_MAX) _conflicts.TryDequeue(out _);
        }

        private async Task<PullResult?> SafePull(long? sinceOverride = null)
        {
            if (!_started || _disposed) return null;
            if (!_pullSem.Wait(0)) return null;
            var task = PullCore(sinceOverride ?? MaxRowVersion);
            try { return await task.ConfigureAwait(false); }
            catch (Exception ex) { PushConflict(Conflict.System("pull", ex.Message)); return null; }
            finally { _pullSem.Release(); }
        }

        private async Task FlushUpsertsCore()
        {
            try
            {
                if (_outbox.IsEmpty) return;

                long deadline = XqlCommon.NowMs() + UPSERT_SLICE_MS;
                do
                {
                    var batch = DrainDedupCells(_outbox, UPSERT_CHUNK);
                    if (batch.Count == 0) break;

                    var resp = await _backend.UpsertCells(batch, _cts.Token).ConfigureAwait(false);

                    if (resp.Errors is { Count: > 0 })
                        foreach (var e in resp.Errors)
                            PushConflict(Conflict.System("upsert", e));

                    if (resp.MaxRowVersion > 0)
                        XqlCommon.InterlockedMax(ref _maxRowVersion, resp.MaxRowVersion);

                    if (resp.Conflicts is { Count: > 0 })
                        foreach (var c in resp.Conflicts) PushConflict(c);

                    // 성공 후 마지막 값 갱신
                    foreach (var e in batch) RememberPushed(Key(e), XqlCommon.Canonicalize(e.Value));
                }
                while (!_outbox.IsEmpty && XqlCommon.NowMs() < deadline);
            }
            catch (Exception ex)
            {
                PushConflict(Conflict.System("upsert.core", ex.Message));
            }
        }

        private async Task<PullResult> PullCore(long sinceVersion)
        {
            var resp = await _backend.PullRows(sinceVersion, _cts.Token).ConfigureAwait(false);

            if (resp.MaxRowVersion > 0)
                XqlCommon.InterlockedMax(ref _maxRowVersion, resp.MaxRowVersion);

            // 서버 패치를 엑셀에 적용 (내부에서 UI 스레드 마샬링)
            if (resp.Patches is { Count: > 0 })
                XqlSheetView.ApplyPlanAndPatches(null, resp.Patches);

            return resp;
        }

        private void OnServerEvent(ServerEvent ev)
        {
            try
            {
                var before = MaxRowVersion;

                if (ev.Patches is { Count: > 0 })
                    XqlSheetView.ApplyPlanAndPatches(null, ev.Patches);

                if (ev.MaxRowVersion > 0)
                    XqlCommon.InterlockedMax(ref _maxRowVersion, ev.MaxRowVersion);

                // 갭이 생기면 보정 Pull
                if (ev.MaxRowVersion > before + 1)
                {
#pragma warning disable CS4014
                    PullSince(before);
#pragma warning restore CS4014
                }
            }
            catch (Exception ex)
            {
                PushConflict(Conflict.System("subscription", ex.Message));
            }
        }

        private static List<EditCell> DrainDedupCells(ConcurrentQueue<EditCell> q, int max)
        {
            var temp = new List<EditCell>(Math.Min(max * 2, 4096));
            for (int i = 0; i < max && q.TryDequeue(out var e); i++) temp.Add(e);
            if (temp.Count <= 1) return temp;

            var map = new Dictionary<string, EditCell>(temp.Count, StringComparer.Ordinal);
            foreach (var e in temp) map[Key(e)] = e; // 마지막 값 우선
            return map.Values.ToList();
        }
    }
}
