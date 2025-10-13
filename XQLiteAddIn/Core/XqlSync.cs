// XqlSync.cs  (ExcelPatchApplier 포함 버전)
using ExcelDna.Integration;
using Newtonsoft.Json;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.Policy;
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
        private volatile bool _pullAgain;
        private long _pullBackoffUntilMs;
        private int _pullErr; // 연속 오류 횟수

        private long _maxRowVersion;
        public long MaxRowVersion => Interlocked.Read(ref _maxRowVersion);

        private readonly Timer _pushTimer;
        private readonly Timer _pullTimer;

        private volatile bool _started;
        private volatile bool _disposed;

        private const int UPSERT_CHUNK = 512;   // 1회 전송 셀 수
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
        private volatile bool _pendingSchemaCheck = false;

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

            // ✅ 구독 시작은 동기 메서드 사용
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

        // ⬇️ 초기화 진입점 (워크북이 열릴 때 한 번 호출)
        public void InitPersistentState(string workbookFullName, string? project = null)
        {
            _workbookFullName = workbookFullName;

            _state = new PersistentState
            {
                Project = project ?? XqlConfig.Project ?? "",
                Workbook = Path.GetFileNameWithoutExtension(workbookFullName) ?? "wb",
            };

            // 워크북에서 K/V 읽기 (UI 스레드에서 안전하게)
            var loaded = new Dictionary<string, string>(StringComparer.Ordinal);
            var done = new ManualResetEventSlim(false);
            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                try
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
                        if (wb != null)
                            loaded = XqlSheet.StateReadAll(wb);
                    }
                    finally { XqlCommon.ReleaseCom(wb); }
                }
                catch { }
                finally { done.Set(); }
            });
            done.Wait(1000); // Excel 바쁘면 그냥 빈 상태로 진행

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
                _pendingSchemaCheck = true;
                _ = Task.Run(async () =>
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
                    finally { _pendingSchemaCheck = false; }
                });
            }
        }

        public void FlushUpsertsNow() => _ = FlushUpsertsNow(false);

        public async Task PullSince(long? sinceOverride = null)
        {
            // 백오프 윈도우엔 스킵
            if (XqlCommon.Monotonic.NowMs() < _pullBackoffUntilMs)
                return;

            if (!await _pullSem.WaitAsync(0).ConfigureAwait(false))
            {
                _pullAgain = true;
                return;
            }
            try
            {
                do
                {
                    _pullAgain = false;
                    try
                    {
                        var since = sinceOverride ?? (_forceFullPull ? 0 : MaxRowVersion);
                        var pr = await _backend.PullRows(since, _cts.Token).ConfigureAwait(false);

                        if (pr.Patches is { Count: > 0 })
                            XqlSheetView.ApplyOnUiThread(pr.Patches);

                        if (pr.MaxRowVersion > 0)
                            XqlCommon.InterlockedMax(ref _maxRowVersion, pr.MaxRowVersion);

                        if (_forceFullPull)
                        {
                            _forceFullPull = false;
                            _state.LastFullPullUtc = DateTime.UtcNow;
                        }
                        _state.LastMaxRowVersion = MaxRowVersion;
                        PersistState();


                        // 스키마 자동 감지(메타가 아직 확인 안 됐으면 1회 더 확인)
                        if (XqlConfig.FullPullWhenSchemaChanged && !_pendingSchemaCheck && string.IsNullOrEmpty(_state.LastSchemaHash))
                        {
                            _pendingSchemaCheck = true;
                            _ = Task.Run(async () =>
                            {
                                try
                                {
                                    var meta = await _backend.TryFetchServerMeta().ConfigureAwait(false);
                                    var hash = meta?["schema_hash"]?.ToString();
                                    if (!string.IsNullOrWhiteSpace(hash))
                                    {
                                        if (!string.Equals(hash, _state.LastSchemaHash, StringComparison.Ordinal))
                                        {
                                            // 다음 사이클에서 보정 Full Pull
                                            _forceFullPull = true;
                                        }
                                        _state.LastSchemaHash = hash;
                                        _state.LastMetaUtc = DateTime.UtcNow;
                                        PersistState();
                                    }
                                }
                                catch { }
                                finally { _pendingSchemaCheck = false; }
                            });
                        }

                        // 성공 → 백오프 해제
                        _pullErr = 0;
                        _pullBackoffUntilMs = 0;
                    }
                    catch (Exception)
                    {
                        // 실패 → 지수 백오프 (최대 8초)
                        _pullErr = Math.Min(_pullErr + 1, 4);
                        _pullBackoffUntilMs = XqlCommon.Monotonic.NowMs() + _pullErr * 2000L;
                        // 실패 시 루프 종료
                        break;
                    }
                }
                while (_pullAgain);
            }
            finally { _pullSem.Release(); }
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

                ExcelAsyncUtil.QueueAsMacro(() =>
                {
                    try
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
                    }
                    catch { }
                });
            }
            catch { }
        }

        private void SafeFlushUpserts()
        {
            if (!_started || _disposed) return;

            // 재진입 방지용 비동기 락(아래 #2 참고)이 있으면 lock 제거 가능
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

        // ⬇️ 교체
        private async Task FlushUpsertsCore()
        {
            try
            {
                if (_outbox.IsEmpty) return;

                long deadline = XqlCommon.Monotonic.NowMs() + UPSERT_SLICE_MS;
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

                    // FlushUpsertsCore 내 성공 후 기록 교체
                    foreach (var e in batch) RememberPushed(Key(e), XqlCommon.Canonicalize(e.Value));
                }
                while (!_outbox.IsEmpty && XqlCommon.Monotonic.NowMs() < deadline);
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

            // ⬇️ 서버 패치를 엑셀에 적용 (UI 스레드 매크로 큐로 안전하게)
            if (resp.Patches is { Count: > 0 })
                XqlSheetView.ApplyOnUiThread(resp.Patches);

            return resp;
        }

        private void OnServerEvent(ServerEvent ev)
        {
            try
            {
                var before = MaxRowVersion;

                if (ev.Patches is { Count: > 0 })
                    XqlSheetView.ApplyOnUiThread(ev.Patches);

                if (ev.MaxRowVersion > 0)
                    XqlCommon.InterlockedMax(ref _maxRowVersion, ev.MaxRowVersion);

                if (ev.MaxRowVersion > before + 1)
                    _ = PullSince(before); // 갭 보정
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
