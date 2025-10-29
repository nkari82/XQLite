// XqlSync.cs  (ExcelPatchApplier 포함 버전, SmartCom 적용)
// ▶ 변경 요약
// 1) PullSince: 빈 시트 또는 since==0 이면 서버 스키마로 [id, ...] 헤더 보장 후 rowsSnapshot 적용
// 2) EnsureHeaderFromServerAsync: 서버 컬럼 조회 → [PK(id) + 나머지] 순서로 헤더 작성
// 3) ApplySnapshotToSheetAsync: 헤더 이름-열 인덱스 매핑으로 스냅샷을 정확히 쓰기(열 밀림 방지)
// 4) 기존 증분 패치/업서트/상태저장 로직은 변경 없음

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

        public void EnqueueIfChanged(string table, string rowKey, string column, object? value)
        {
            // PK(id)는 서버가 관리: 사용자의 직접 편집은 무시
            if (!string.IsNullOrEmpty(column) && column.Equals("id", StringComparison.OrdinalIgnoreCase))
                return;

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

            // 워크북에서 K/V 읽기 (UI 스레드에서 안전하게, SmartCom 사용)
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
                                    wbHit = cur;  // 소유권 유지(반환 전까지 유지)
                                    cur = null;   // 현 지역 참조는 방출
                                    break;
                                }
                            }
                            finally
                            {
                                // cur가 hit가 아니면 SmartCom이 아니므로 명시 해제 불가 → GC에 맡김
                                // (여기서는 곧바로 루프 종료/반환하므로 추가 참조 없음)
                            }
                        }

                        if (wbHit == null)
                            wbHit = appW.Value.ActiveWorkbook;

                        return wbHit != null
                            ? XqlSheet.StateReadAll(wbHit)
                            : new Dictionary<string, string>(StringComparer.Ordinal);
                    }
                    finally
                    {
                        // wbHit는 XqlSheet.StateReadAll 내부에서 필요한 범위만 접근 후 RCW 유지 필요 없음.
                        // 별도 Release 없이 범위를 벗어나면 RCW는 GC로 회수됨(현 스코프에서 추가 참조 없음).
                    }
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
                        if (!string.IsNullOrWhiteSpace(hash) &&
                            !string.Equals(hash, _state.LastSchemaHash, StringComparison.Ordinal))
                            _forceFullPull = true;

                        _state.LastSchemaHash = hash;
                        _state.LastMetaUtc = DateTime.UtcNow;
                        PersistState();
                    }
                    catch { /* ignore */ }
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
                // ── 부트스트랩 판단: 시트가 비었거나, since==0 이거나, 강제 플래그
                var ui = await OnExcelThreadAsync(() =>
                {
                    var app = ExcelDnaUtil.Application as Excel.Application;
                    using var wsW = SmartCom<Excel.Worksheet>.Wrap(app?.ActiveSheet);
                    if (wsW?.Value == null) return (needs: false, sheet: null as Excel.Worksheet, table: "", key: "");

                    var sm = _sheet.GetOrCreateSheet(wsW.Value.Name);
                    var table = string.IsNullOrWhiteSpace(sm.TableName) ? wsW.Value.Name : sm.TableName!;
                    var key = string.IsNullOrWhiteSpace(sm.KeyColumn) ? "id" : sm.KeyColumn!;
                    bool needs = XqlSheet.NeedsBootstrap(wsW.Value);
                    return (needs: needs, sheet: wsW.Value, table, key);
                }).ConfigureAwait(false);

                var since = sinceOverride ?? (_forceFullPull ? 0 : MaxRowVersion);

                if (ui.sheet != null && (ui.needs || since == 0 || MaxRowVersion == 0))
                {
                    // 1) 서버 스키마로 헤더 보장
                    var header = await EnsureHeaderFromServerAsync(ui.sheet!, ui.table, ui.key).ConfigureAwait(false);

                    // 2) rowsSnapshot 받아 한 방에 쓰기
                    var snap = await _backend.FetchRowsSnapshot(ui.table, _cts.Token).ConfigureAwait(false);
                    await ApplySnapshotToSheetAsync(ui.sheet!, ui.table, header, snap).ConfigureAwait(false);

                    // 3) 상태/캐시
                    XqlSheetView.InvalidateHeaderCache(ui.sheet!.Name);
                    _pullErr = 0;
                    _forceFullPull = false;
                    return;
                }

                // ── 기존 증분 경로
                var pr = await _backend.PullRows(since, _cts.Token).ConfigureAwait(false);

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

        // Private ============================================================

        // (신규) 서버 스키마 기반으로 헤더를 [PK(id) + 기타] 순으로 보장
        private async Task<List<string>> EnsureHeaderFromServerAsync(Excel.Worksheet ws, string table, string key)
        {
            var cols = await _backend.GetTableColumns(table, _cts.Token).ConfigureAwait(false);
            if (cols.Count == 0)
            {
                // 서버에 테이블이 아직 없을 수 있음 → 생성 후 재질의
                await _backend.TryCreateTable(table, key, _cts.Token).ConfigureAwait(false);
                cols = await _backend.GetTableColumns(table, _cts.Token).ConfigureAwait(false);
            }

            // 순서: PK 먼저, 나머지는 서버 순서 유지
            string pk = cols.FirstOrDefault(c => c.pk)?.name ?? key;
            var ordered = new List<string> { pk };
            ordered.AddRange(cols.Where(c => !c.pk).Select(c => c.name));

            // 현재 헤더와 다르면 교체
            await OnExcelThreadAsync(() =>
            {
                var (hdrRange, names0) = XqlSheet.GetHeaderAndNames(ws);
                using var hdr = SmartCom<Excel.Range>.Wrap(hdrRange ?? XqlSheet.GetHeaderRange(ws));
                if (hdr?.Value == null) return 0;

                bool same = names0.Count == ordered.Count &&
                            names0.Zip(ordered, (a, b) => string.Equals(a, b, StringComparison.OrdinalIgnoreCase)).All(x => x);
                if (same) return 0;

                for (int i = 0; i < ordered.Count; i++)
                {
                    using var c = SmartCom<Excel.Range>.Acquire(() => (Excel.Range)ws.Cells[hdr.Value.Row, hdr.Value.Column + i]);
                    try { if (c?.Value != null) c.Value!.Value2 = ordered[i]; } catch { }
                }
                // 남은 이전 헤더 비우기
                for (int j = ordered.Count; j < Math.Max(ordered.Count, names0.Count); j++)
                {
                    using var c = SmartCom<Excel.Range>.Acquire(() => (Excel.Range)ws.Cells[hdr.Value.Row, hdr.Value.Column + j]);
                    try { if (c?.Value != null) c.Value!.Value2 = ""; } catch { }
                }
                return 0;
            }).ConfigureAwait(false);

            // 메타 레지스트리에도 반영
            _sheet.EnsureColumns(ws.Name, ordered);
            return ordered;
        }

        // (신규) rowsSnapshot을 헤더 매핑대로 시트에 씀
        private async Task ApplySnapshotToSheetAsync(Excel.Worksheet ws, string table, List<string> header, List<RowPatch> rows)
        {
            if (rows == null) rows = new List<RowPatch>();

            await OnExcelThreadAsync(() =>
            {
                var (hdrRange, _) = XqlSheet.GetHeaderAndNames(ws);
                using var hdr = SmartCom<Excel.Range>.Wrap(hdrRange ?? XqlSheet.GetHeaderRange(ws));
                if (hdr?.Value == null) return 0;

                int firstDataRow = hdr.Value.Row + 1;
                int colCount = header.Count;

                // 데이터 영역 Clear
                using (var rangeAll = SmartCom<Excel.Range>.Acquire(() =>
                    (Excel.Range)ws.Range[
                        ws.Cells[firstDataRow, hdr.Value.Column],
                        ws.Cells[ws.Rows.Count, hdr.Value.Column + colCount - 1]
                    ]))
                { try { if (rangeAll?.Value != null) rangeAll.Value!.ClearContents(); } catch { } }

                // 헤더 인덱스 맵
                var idx = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                for (int i = 0; i < header.Count; i++) idx[header[i]] = i;

                int r = firstDataRow;
                foreach (var p in rows)
                {
                    // 우선 id/row_key
                    object? id = p.RowKey;
                    if (id == null && p.Cells.TryGetValue("id", out var id2)) id = id2;

                    if (id != null && idx.TryGetValue("id", out int idCol0))
                    {
                        using var c = SmartCom<Excel.Range>.Acquire(() => (Excel.Range)ws.Cells[r, hdr.Value.Column + idCol0]);
                        try { if (c?.Value != null) c.Value!.Value2 = id; } catch { }
                    }

                    foreach (var kv in p.Cells)
                    {
                        if (!idx.TryGetValue(kv.Key, out int ci)) continue;
                        using var c = SmartCom<Excel.Range>.Acquire(() => (Excel.Range)ws.Cells[r, hdr.Value.Column + ci]);
                        try { if (c?.Value != null) c.Value!.Value2 = kv.Value is null ? null : kv.Value; } catch { }
                    }
                    r++;
                }

                if (hdr.Value != null)
                {
                    var sm = _sheet.GetOrCreateSheet(ws.Name);
                    XqlSheetView.ApplyHeaderUi(ws, hdr.Value, sm, withValidation: true);
                    XqlSheet.SetHeaderMarker(ws, hdr.Value);
                    XqlSheetView.RegisterTableSheet(table, ws.Name);
                }
                return 0;
            }).ConfigureAwait(false);
        }

        // (기존) 풀 전체가 처음이거나, 증분이 0건인 초기 상태에서 호출 — 사용하지 않지만 남겨둠(호환)
        private async Task ApplyBootstrapAsync(PullResult pr)
        {
            var meta = await _backend.TryFetchServerMeta().ConfigureAwait(false);
            if (meta == null) return;

            var schema = new Dictionary<string, List<string>>(StringComparer.Ordinal);

            if (meta["tables"] is JArray tablesNew)
            {
                foreach (var t in tablesNew.OfType<JObject>())
                {
                    var tname = (t["name"] ?? t["table_name"])?.ToString();
                    if (string.IsNullOrWhiteSpace(tname)) continue;

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

                // UI 스레드에서만 상태 기록 (SmartCom 사용)
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
                                    wbHit = cur;
                                    cur = null;
                                    break;
                                }
                            }
                            finally
                            {
                                // cur 로컬 참조는 여기서만 사용되었고 루프 계속 → 추가 처리 불필요
                            }
                        }

                        wbHit ??= appW.Value.ActiveWorkbook;
                        if (wbHit != null)
                            XqlSheet.StateSetMany(wbHit, kv);
                    }
                    finally
                    {
                        // wbHit 참조는 StateSetMany 내에서만 사용되며, 이 스코프에서 더 이상 보유하지 않음
                    }
                    return 0;
                });
            }
            catch { /* ignore */ }
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
