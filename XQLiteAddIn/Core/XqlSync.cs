// XqlSync.cs  (pending/재큐/스냅샷 안정화 + 자동부트스트랩 통합판)
// - 배치 전송 경합 차단(_pendingKeys)
// - 스냅샷 적용시 헤더명→열 인덱스 매핑으로 정확 기입
// - 성공시에만 LRU 갱신, 실패/예외는 전량 재큐
// - PK(id) 클라이언트 편집 무시
// - 시작/타이머/구독 이벤트의 자동 Pull도 항상 PullSince(…)=부트스트랩 경로 공유

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
using static XQLite.AddIn.XqlCommon;
using Microsoft.Office.Interop.Excel;

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

        // 편집 아웃박스(전송 전 임시 보관)
        private readonly ConcurrentQueue<EditCell> _outbox = new();

        // 전송 중 키(중복 전송/경합 방지)
        private readonly ConcurrentDictionary<string, byte> _pendingKeys = new(StringComparer.Ordinal);

        private readonly SemaphoreSlim _pushSem = new(1, 1);
        private readonly SemaphoreSlim _pullSem = new(1, 1);
        private int _pulling; // 0/1
        public bool IsPulling => System.Threading.Volatile.Read(ref _pulling) == 1;
        public event Action<bool>? PullStateChanged;

        private long _pullBackoffUntilMs;
        private int _pullErr;

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
            _backend = backend ?? throw new ArgumentNullException(nameof(backend));

            _pushIntervalMs = Math.Max(250, pushIntervalMs);
            _pullIntervalMs = Math.Max(1000, pullIntervalMs);

            _pushTimer = new Timer(_ => SafeFlushUpserts(), null, Timeout.Infinite, Timeout.Infinite);
            _pullTimer = new Timer(_ => _ = SafePull(), null, Timeout.Infinite, Timeout.Infinite);
        }

        // ─────────────────────────────────── 공용 API

        public void Start()
        {
            if (_disposed || _started) return;
            _started = true;

            _cts = new CancellationTokenSource();

            _pushTimer.Change(_pushIntervalMs, _pushIntervalMs);
            _pullTimer.Change(_pullIntervalMs, _pullIntervalMs);

            _backend.StartSubscription(OnServerEvent, MaxRowVersion);

            // ★ 시작 직후 1회: 항상 부트스트랩 경로를 타도록 0부터 Pull
            _ = Task.Run(async () =>
            {
                try
                {
                    await Task.Delay(200).ConfigureAwait(false); // Excel 초기화 안정화 버퍼
                    await PullSince(0).ConfigureAwait(false);
                }
                catch { /* ignore */ }
            });
        }

        public void Stop()
        {
            if (!_started) return;
            _started = false;

            try { _cts.Cancel(); } catch { }

            try { _pushTimer.Change(Timeout.Infinite, Timeout.Infinite); } catch { }
            try { _pullTimer.Change(Timeout.Infinite, Timeout.Infinite); } catch { }

            try { _backend.StopSubscription(); } catch { }
        }

        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;
            try { Stop(); } catch { }
            try { _pushTimer.Dispose(); } catch { }
            try { _pullTimer.Dispose(); } catch { }
            try { _cts.Dispose(); } catch { }
        }

        private static string Key(EditCell e) => $"{e.Table}\n{XqlCommon.ValueToString(e.RowKey)}\n{e.Column}";

        /// <summary>사용자 편집 enqueue(동일 값 연속 전송 방지). id는 클라이언트 편집 무시.</summary>
        public void EnqueueIfChanged(string table, string rowKey, string column, object? value)
        {
            if (!string.IsNullOrEmpty(column) && column.Equals("id", StringComparison.OrdinalIgnoreCase))
                return; // PK는 서버 관리

            var e = new EditCell(Table: table, RowKey: rowKey, Column: column, Value: value);
            var k = Key(e);
            var norm = XqlCommon.Canonicalize(value);
            if (IsSameAsLast(k, norm)) return; // 직전 성공 값과 같으면 전송 불필요

            _outbox.Enqueue(e);
        }

        public bool TryDequeueConflict(out Conflict c) => _conflicts.TryDequeue(out c);

        /// <summary>워크북 열릴 때 1회 상태 로드</summary>
        public void InitPersistentState(string workbookFullName, string? project = null)
        {
            _workbookFullName = workbookFullName;
            var wbName = Path.GetFileNameWithoutExtension(workbookFullName) ?? "wb";

            var proj = (project ?? XqlConfig.Project ?? "").Trim();
            if (string.IsNullOrEmpty(proj)) proj = wbName;

            _state = new PersistentState
            {
                Project = proj,
                Workbook = wbName,
                LastMaxRowVersion = 0,
                LastFullPullUtc = DateTime.MinValue
            };

            // 워크북 K/V 읽기 (UI 스레드)
            Dictionary<string, string> loaded;
            try
            {
                loaded = OnExcelThreadAsync(() =>
                {
                    using var appW = SmartCom<Excel.Application>.Wrap((Excel.Application)ExcelDnaUtil.Application);
                    if (appW.Value == null) return new Dictionary<string, string>(StringComparer.Ordinal);

                    using var booksW = SmartCom<Workbooks>.Wrap(appW.Value.Workbooks);
                    Excel.Workbook? wbHit = null;

                    try
                    {
                        int count = booksW.Value?.Count ?? 0;
                        for (int i = 1; i <= count; i++)
                        {
                            Excel.Workbook? cur = null;
                            try
                            {
                                cur = booksW.Value![i];
                                if (string.Equals(cur.FullName, workbookFullName, StringComparison.OrdinalIgnoreCase))
                                {
                                    wbHit = cur; cur = null; break;
                                }
                            }
                            finally { }
                        }

                        wbHit ??= appW.Value.ActiveWorkbook;
                        return wbHit != null ? XqlSheet.StateReadAll(wbHit)
                                             : new Dictionary<string, string>(StringComparer.Ordinal);
                    }
                    finally { }
                }).GetAwaiter().GetResult();
            }
            catch
            {
                loaded = new Dictionary<string, string>(StringComparer.Ordinal);
            }

            if (loaded.TryGetValue("last_max_row_version", out var s) && long.TryParse(s, out var l))
                _state.LastMaxRowVersion = l;
            if (loaded.TryGetValue("last_schema_hash", out var h)) _state.LastSchemaHash = h;
            if (loaded.TryGetValue("last_full_pull_utc", out var f) && DateTime.TryParse(f, out var dt))
                _state.LastFullPullUtc = dt;

            _forceFullPull = XqlConfig.AlwaysFullPullOnStartup;
            _state.LastSessionId = Guid.NewGuid().ToString("N");
            PersistState();

            if (XqlConfig.FullPullWhenSchemaChanged)
            {
                Task.Run(async () =>
                {
                    try
                    {
                        var meta = await _backend.TryFetchServerMeta().ConfigureAwait(false);
                        var hash = meta?["schema_hash"]?.ToString();
                        if (!string.IsNullOrWhiteSpace(hash) &&
                            !string.Equals(hash, _state.LastSchemaHash, StringComparison.Ordinal))
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
                // Excel 스레드에서 "문자열 스냅샷"만 수집 (RCW 금지)
                var ui = await OnExcelThreadAsync(() =>
                {
                    var app = ExcelDnaUtil.Application as Excel.Application;
                    var ws = app?.ActiveSheet as Excel.Worksheet;
                    if (ws == null) return (needs: false, sheet: "", table: "", key: "");

                    var sm = _sheet.GetOrCreateSheet(ws.Name);
                    var table = string.IsNullOrWhiteSpace(sm.TableName) ? ws.Name : sm.TableName!;
                    var key = string.IsNullOrWhiteSpace(sm.KeyColumn) ? "id" : sm.KeyColumn!;
                    bool needs = XqlSheet.NeedsBootstrap(ws);
                    return (needs, sheet: ws.Name, table, key);
                }).ConfigureAwait(false);

                var since = sinceOverride ?? (_forceFullPull ? 0 : MaxRowVersion);

                // 부트스트랩(또는 강제 full pull)
                if (!string.IsNullOrEmpty(ui.sheet) && (ui.needs || since == 0 || MaxRowVersion == 0))
                {
                    var header = await EnsureHeaderFromServerAsync(ui.sheet, ui.table, ui.key).ConfigureAwait(false);

                    var snap = await _backend.FetchRowsSnapshot(ui.table, _cts.Token).ConfigureAwait(false);
                    await ApplySnapshotToSheetAsync(ui.sheet, ui.table, header, snap).ConfigureAwait(false);

                    XqlSheetView.InvalidateHeaderCache(ui.sheet);
                    _pullErr = 0;
                    _forceFullPull = false;
                    return;
                }

                // 증분
                var pr = await _backend.PullRows(since, _cts.Token).ConfigureAwait(false);
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
            if (!force && (!_started)) return;

            if (!_pushSem.Wait(0)) return;
            try { await FlushUpsertsCore().ConfigureAwait(false); }
            finally { _pushSem.Release(); }
        }

        // ─────────────────────────────────── 내부 로직

        // 서버 스키마로 헤더 보장: [PK(id) + 기타(서버 순서 유지)]
        // 시그니처 변경: ws 대신 sheetName
        private async Task<List<string>> EnsureHeaderFromServerAsync(string sheetName, string table, string key)
        {
            var cols = await _backend.GetTableColumns(table, _cts.Token).ConfigureAwait(false);
            if (cols.Count == 0)
            {
                await _backend.TryCreateTable(table, key, _cts.Token).ConfigureAwait(false);
                cols = await _backend.GetTableColumns(table, _cts.Token).ConfigureAwait(false);
            }

            string pk = cols.FirstOrDefault(c => c.pk)?.name ?? key;
            var ordered = new List<string> { pk };
            ordered.AddRange(cols.Where(c => !c.pk).Select(c => c.name));

            await OnExcelThreadAsync(() =>
            {
                var app = (Excel.Application)ExcelDnaUtil.Application;

                // 여기서 RCW 재획득 (이름 기반)
                using var wsW = SmartCom<Excel.Worksheet>.Wrap(XqlSheet.FindWorksheet(app, sheetName));
                if (wsW?.Value == null) return 0;

                // 기존 헤더/UsedRange 등 "읽기"는 안전하게 수행 (동일성 비교는 Optional)
                var (hdrRange, names0) = XqlSheet.GetHeaderAndNames(wsW.Value);
                using var hdr = SmartCom<Excel.Range>.Wrap(hdrRange ?? XqlSheet.GetHeaderRange(wsW.Value));
                if (hdr?.Value == null) return 0;

                bool same = names0.Count == ordered.Count &&
                            names0.Zip(ordered, (a, b) => string.Equals(a, b, StringComparison.OrdinalIgnoreCase)).All(x => x);
                if (!same)
                {
                    for (int i = 0; i < ordered.Count; i++)
                    {
                        using var c = SmartCom<Excel.Range>.Acquire(() => (Excel.Range)wsW.Value.Cells[hdr.Value.Row, hdr.Value.Column + i]);
                        try { if (c?.Value != null) c.Value!.Value2 = ordered[i]; } catch { }
                    }
                    // 남은 칸 정리
                    for (int j = ordered.Count; j < Math.Max(ordered.Count, names0.Count); j++)
                    {
                        using var c = SmartCom<Excel.Range>.Acquire(() => (Excel.Range)wsW.Value.Cells[hdr.Value.Row, hdr.Value.Column + j]);
                        try { if (c?.Value != null) c.Value!.Value2 = ""; } catch { }
                    }
                }

                // 마커 보장
                try { XqlSheet.SetHeaderMarker(wsW.Value, hdr.Value); } catch { }

                return 0;
            }).ConfigureAwait(false);

            // 메타 레지스트리 갱신도 문자열 기반
            try { _sheet.EnsureColumns(sheetName, ordered); } catch { }

            return ordered;
        }

        // rowsSnapshot을 헤더 매핑대로 시트에 씀(열 밀림 방지)
        private async Task ApplySnapshotToSheetAsync(string sheetName, string table, List<string> header, List<RowPatch> rows)
        {
            rows ??= new List<RowPatch>();

            await OnExcelThreadAsync(() =>
            {
                var app = (Excel.Application)ExcelDnaUtil.Application;
                using var wsW = SmartCom<Excel.Worksheet>.Wrap(XqlSheet.FindWorksheet(app, sheetName));
                if (wsW?.Value == null) return 0;

                var (hdrRange, _) = XqlSheet.GetHeaderAndNames(wsW.Value);
                using var hdr = SmartCom<Excel.Range>.Wrap(hdrRange ?? XqlSheet.GetHeaderRange(wsW.Value));
                if (hdr?.Value == null) return 0;

                int firstDataRow = hdr.Value.Row + 1;
                int colCount = header.Count;

                // 데이터 영역 Clear
                using (var rangeAll = SmartCom<Excel.Range>.Acquire(() =>
                    (Excel.Range)wsW.Value.Range[
                        wsW.Value.Cells[firstDataRow, hdr.Value.Column],
                        wsW.Value.Cells[wsW.Value.Rows.Count, hdr.Value.Column + colCount - 1]
                    ]))
                { try { if (rangeAll?.Value != null) rangeAll.Value!.ClearContents(); } catch { } }

                // 헤더 인덱스
                var idx = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                for (int i = 0; i < header.Count; i++) idx[header[i]] = i;

                int r = firstDataRow;
                foreach (var p in rows)
                {
                    object? id = p.RowKey;
                    if (id == null && p.Cells.TryGetValue("id", out var id2)) id = id2;

                    if (id != null && idx.TryGetValue("id", out int idCol0))
                    {
                        using var c = SmartCom<Excel.Range>.Acquire(() => (Excel.Range)wsW.Value.Cells[r, hdr.Value.Column + idCol0]);
                        try { if (c?.Value != null) c.Value!.Value2 = id; } catch { }
                    }

                    foreach (var kv in p.Cells)
                    {
                        if (!idx.TryGetValue(kv.Key, out int ci)) continue;
                        using var c = SmartCom<Excel.Range>.Acquire(() => (Excel.Range)wsW.Value.Cells[r, hdr.Value.Column + ci]);
                        try { if (c?.Value != null) c.Value!.Value2 = kv.Value is null ? null : kv.Value; } catch { }
                    }
                    r++;
                }

                var sm = _sheet.GetOrCreateSheet(wsW.Value.Name);
                XqlSheetView.ApplyHeaderUi(wsW.Value, hdr.Value, sm, withValidation: true);
                XqlSheet.SetHeaderMarker(wsW.Value, hdr.Value);
                XqlSheetView.RegisterTableSheet(table, wsW.Value.Name);

                return 0;
            }).ConfigureAwait(false);
        }

        // 풀 결과 패치 적용(UI 스레드 마샬링)
        private Task ApplyIncrementalPatches(PullResult pr)
        {
            if (pr?.Patches is { Count: > 0 })
                XqlSheetView.ApplyPlanAndPatches(null, pr.Patches);
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

                _ = XqlCommon.OnExcelThreadAsync(() =>
                {
                    using var appW = SmartCom<Excel.Application>.Wrap((Excel.Application)ExcelDnaUtil.Application);
                    if (appW.Value == null) return 0;

                    using var booksW = SmartCom<Workbooks>.Wrap(appW.Value.Workbooks);
                    Excel.Workbook? wbHit = null;

                    try
                    {
                        int count = booksW.Value?.Count ?? 0;
                        for (int i = 1; i <= count; i++)
                        {
                            Excel.Workbook? cur = null;
                            try
                            {
                                cur = booksW.Value![i];
                                if (string.Equals(cur.FullName, _workbookFullName, StringComparison.OrdinalIgnoreCase))
                                {
                                    wbHit = cur; cur = null; break;
                                }
                            }
                            finally { }
                        }

                        wbHit ??= appW.Value.ActiveWorkbook;
                        if (wbHit != null) XqlSheet.StateSetMany(wbHit, kv);
                    }
                    finally { }
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

        // ★ 자동 Pull 경로도 항상 PullSince(…)=부트스트랩 공유
        private async Task<PullResult?> SafePull(long? sinceOverride = null)
        {
            if (!_started || _disposed) return null;
            try
            {
                await PullSince(sinceOverride).ConfigureAwait(false);
                return null;
            }
            catch (Exception ex)
            {
                PushConflict(Conflict.System("pull", ex.Message));
                return null;
            }
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

                    // 전송 전에 pending 마킹
                    var keys = new List<string>(batch.Count);
                    foreach (var e in batch)
                    {
                        var k = Key(e);
                        keys.Add(k);
                        _pendingKeys.TryAdd(k, 1);
                    }

                    bool success = false;
                    UpsertResult? resp = null;

                    try
                    {
                        resp = await _backend.UpsertCells(batch, _cts.Token).ConfigureAwait(false);

                        var hasErrors = resp?.Errors is { Count: > 0 };
                        var hasConflicts = resp?.Conflicts is { Count: > 0 };
                        success = !(hasErrors || hasConflicts);

                        if (hasErrors)
                            foreach (var e in resp!.Errors!)
                                PushConflict(Conflict.System("upsert", e));

                        if (hasConflicts)
                            foreach (var c in resp!.Conflicts!)
                                PushConflict(c);

                        if (resp?.MaxRowVersion > 0)
                            XqlCommon.InterlockedMax(ref _maxRowVersion, resp.MaxRowVersion);
                    }
                    catch (OperationCanceledException)
                    {
                        Requeue(_outbox, batch);
                        return;
                    }
                    catch (Exception ex)
                    {
                        Requeue(_outbox, batch);
                        PushConflict(Conflict.System("upsert.core", ex.Message));
                        success = false;
                    }
                    finally
                    {
                        // pending 해제 (성공/실패 공통)
                        foreach (var k in keys) _pendingKeys.TryRemove(k, out _);
                    }

                    if (!success)
                    {
                        // 보수적: 실패/컨플릭트 시 전체 재큐
                        Requeue(_outbox, batch);
                        break;
                    }

                    // 성공: LRU 갱신
                    foreach (var e in batch)
                        RememberPushed(Key(e), XqlCommon.Canonicalize(e.Value));

                } while (!_outbox.IsEmpty && XqlCommon.NowMs() < deadline);
            }
            catch (Exception ex)
            {
                PushConflict(Conflict.System("upsert.core.outer", ex.Message));
            }
        }

        private async Task<PullResult> PullCore(long sinceVersion)
        {
            var resp = await _backend.PullRows(sinceVersion, _cts.Token).ConfigureAwait(false);

            if (resp.MaxRowVersion > 0)
                XqlCommon.InterlockedMax(ref _maxRowVersion, resp.MaxRowVersion);

            if (resp.Patches is { Count: > 0 })
                XqlSheetView.ApplyPlanAndPatches(null, resp.Patches);

            return resp;
        }

        private void OnServerEvent(ServerEvent ev)
        {
            try
            {
                var before = MaxRowVersion;

                // ★ 시트가 부트스트랩 필요하면 패치 대신 풀 부트스트랩
                _ = Task.Run(async () =>
                {
                    try
                    {
                        bool needs = await OnExcelThreadAsync(() =>
                        {
                            var app = ExcelDnaUtil.Application as Excel.Application;
                            using var wsW = SmartCom<Excel.Worksheet>.Wrap(app?.ActiveSheet);
                            return wsW?.Value != null && XqlSheet.NeedsBootstrap(wsW.Value);
                        }).ConfigureAwait(false);

                        if (needs)
                        {
                            await PullSince(0).ConfigureAwait(false);
                        }
                        else
                        {
                            if (ev.Patches is { Count: > 0 })
                                XqlSheetView.ApplyPlanAndPatches(null, ev.Patches);
                        }
                    }
                    catch (Exception ex) { PushConflict(Conflict.System("subscription", ex.Message)); }
                });

                if (ev.MaxRowVersion > 0)
                    XqlCommon.InterlockedMax(ref _maxRowVersion, ev.MaxRowVersion);

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

        // ─────────────────── 배치 빌드/재큐 헬퍼

        private static void Requeue(ConcurrentQueue<EditCell> q, List<EditCell> items)
        {
            // 최신 편집이 뒤로 밀리는 정도는 허용(유실 방지가 더 중요)
            foreach (var e in items)
                q.Enqueue(e);
        }

        /// <summary>
        /// 큐에서 최대 max개를 꺼내면서
        /// 1) 전송 중인 키는 제외하고 지연 재큐
        /// 2) 같은 키는 마지막(최신) 값만 남김
        /// </summary>
        private List<EditCell> DrainDedupCells(ConcurrentQueue<EditCell> q, int max)
        {
            var temp = new List<EditCell>(Math.Min(max * 2, 4096));
            for (int i = 0; i < max && q.TryDequeue(out var e); i++) temp.Add(e);
            if (temp.Count == 0) return temp;

            var defer = new List<EditCell>(temp.Count);
            var map = new Dictionary<string, EditCell>(temp.Count, StringComparer.Ordinal);

            foreach (var e in temp)
            {
                var k = Key(e);
                if (_pendingKeys.ContainsKey(k))
                {
                    defer.Add(e); // 전송중: 이번 배치에서 제외
                    continue;
                }
                map[k] = e; // 최신 값 우선
            }

            // 제외분은 지연 재큐(유실 방지)
            if (defer.Count > 0)
                Requeue(q, defer);

            return map.Values.ToList();
        }
    }
}
