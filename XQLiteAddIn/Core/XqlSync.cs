// XqlSync.cs  (ExcelPatchApplier 포함 버전)
using EnvDTE;
using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;
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
using XQLite.AddIn;
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

        // XqlSync.cs

        // ✅ 기존 PullSince를 아래로 완전히 교체
        public async Task PullSince(long? sinceOverride = null)
        {
            // 백오프 윈도우면 스킵
            if (XqlCommon.Monotonic.NowMs() < _pullBackoffUntilMs)
                return;


            // 이미 실행 중이면 무시
            if (System.Threading.Interlocked.Exchange(ref _pulling, 1) == 1)
                return;
            PullStateChanged?.Invoke(true);
            if (!await _pullSem.WaitAsync(0).ConfigureAwait(false))
            {
                // 세마포어도 잡혔으면 바로 종료 처리
                System.Threading.Interlocked.Exchange(ref _pulling, 0);
                PullStateChanged?.Invoke(false);
                return;
            }

            try
            {

                try
                {
                    // 1) 기본은 증분, 강제 풀이면 0
                    var since = sinceOverride ?? (_forceFullPull ? 0 : MaxRowVersion);
                    var pr = await _backend.PullRows(since, _cts.Token).ConfigureAwait(false);

                    // ⬇️ ① FULL(=0) 이고, 패치가 있으면 "부트스트랩 적용" 시도
                    if (since == 0 && pr.Patches != null && pr.Patches.Count > 0)
                    {
                        await ApplyBootstrapSnapshot(pr).ConfigureAwait(false);
                        _pullErr = 0; // 성공으로 간주
                        Interlocked.Exchange(ref _maxRowVersion, pr.MaxRowVersion);
                        PersistState();
                        return; // 부트스트랩 경로 종료
                    }

                    // (기존) 증분 적용 경로 …
                    await ApplyIncrementalPatches(pr).ConfigureAwait(false);

                    // 2) 패치가 0건이고, 아직 로컬 버전이 0(=초기 워크북)이라면 → 서버에서 Full Pull 재시도
                    if ((pr.Patches == null || pr.Patches.Count == 0) &&
                        since != 0 &&
                        MaxRowVersion == 0 &&
                        _state.LastMaxRowVersion == 0)
                    {
                        pr = await _backend.PullRows(0, _cts.Token).ConfigureAwait(false);

                        // 3) 패치가 있으면 UI 반영
                        await ApplyIncrementalPatches(pr).ConfigureAwait(false);
                    }

                    // 4) 버전 반영
                    if (pr.MaxRowVersion > 0)
                        XqlCommon.InterlockedMax(ref _maxRowVersion, pr.MaxRowVersion);

                    if (_forceFullPull)
                    {
                        _forceFullPull = false;
                        _state.LastFullPullUtc = DateTime.UtcNow;
                    }
                    _state.LastMaxRowVersion = MaxRowVersion;
                    PersistState();

                    // (옵션) 스키마 변경 감지 작업은 기존 로직 유지
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
                                        _forceFullPull = true;
                                    _state.LastSchemaHash = hash;
                                    _state.LastMetaUtc = DateTime.UtcNow;
                                    PersistState();
                                }
                            }
                            catch { }
                            finally { _pendingSchemaCheck = false; }
                        });
                    }

                    // 성공 → 백오프 초기화
                    _pullErr = 0;
                    _pullBackoffUntilMs = 0;
                }
                catch
                {
                    // 실패 → 지수 백오프(최대 8초)
                    _pullErr = Math.Min(_pullErr + 1, 4);
                    _pullBackoffUntilMs = XqlCommon.Monotonic.NowMs() + _pullErr * 2000L;
                }
            }
            finally
            {
                _pullSem.Release();
                System.Threading.Interlocked.Exchange(ref _pulling, 0);
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

        // FULL Pull 전용: 메타헤더가 없으면 만들고, 시트를 초기화한 뒤 스냅샷을 채워 넣는다.
        private async Task ApplyBootstrapSnapshot(PullResult pr)
        {
            var app = (Excel.Application)ExcelDnaUtil.Application;
            await Task.Yield(); // UI 양보

            using var scope = new XqlCommon.ExcelBatchScope(app);
            foreach (var grp in pr.Patches.GroupBy(p => p.Table, StringComparer.Ordinal))
            {
                string table = grp.Key ?? "Sheet1";
                Excel.Worksheet? ws = null;
                Excel.Range? header = null;
                try
                {
                    // 1) 대상 시트 얻기(없으면 생성)
                    ws = XqlSheet.FindWorksheet(app, table) ?? (Excel.Worksheet)app.Worksheets.Add();
                    if (!string.Equals(ws.Name, table, StringComparison.Ordinal)) ws.Name = table;

                    // 2) 헤더 이름 구성: 첫 패치의 cells 키들로 결정 (정렬 안정성 확보)
                    var firstCells = grp.FirstOrDefault()?.Cells ?? new Dictionary<string, object?>();
                    var colNames = firstCells.Keys
                                             .Where(k => !string.Equals(k, "id", StringComparison.OrdinalIgnoreCase)) // id는 맨 앞으로
                                             .OrderBy(k => k, StringComparer.Ordinal)
                                             .ToList();
                    // id, row_version, updated_at, deleted 메타는 항상 보장(앞 쪽)
                    var headerNames = new List<string> { "id", "row_version", "updated_at", "deleted" };
                    foreach (var c in colNames)
                        if (!headerNames.Contains(c, StringComparer.OrdinalIgnoreCase))
                            headerNames.Add(c);

                    // 3) 시트 초기화(헤더 + 본문)
                    ws.Cells.Clear();
                    for (int i = 0; i < headerNames.Count; i++)
                        (ws.Cells[1, i + 1] as Excel.Range)!.Value2 = headerNames[i];

                    header = XqlSheet.GetHeaderRange(ws); // 1행 전체
                                                          // 메타 등록 + UI(툴팁/검증)
                    var sm = XqlAddIn.Sheet!.GetOrCreateSheet(ws.Name);
                    XqlAddIn.Sheet!.EnsureColumns(ws.Name, headerNames);
                    XqlSheetView.ApplyHeaderUi(ws, header, sm, withValidation: true);
                    XqlSheet.SetHeaderMarker(ws, header); // 마커 박제

                    // 4) 본문 채우기(행당 id 열은 필수로 사용)
                    int row = 2;
                    foreach (var p in grp.OrderBy<RowPatch, object>(x => x.RowKey, Comparer<object>.Default))
                    {
                        if (p.Deleted) continue;
                        var cells = p.Cells ?? new Dictionary<string, object?>();

                        (ws.Cells[row, 1] as Excel.Range)!.Value2 = p.RowKey; // id
                                                                              // meta 기본값
                        (ws.Cells[row, 2] as Excel.Range)!.Value2 = p.RowVersion;    // row_version
                        (ws.Cells[row, 3] as Excel.Range)!.Value2 = DateTime.Now;    // updated_at (표시용)
                        (ws.Cells[row, 4] as Excel.Range)!.Value2 = 0;               // deleted

                        // 나머지 데이터
                        for (int c = 5; c <= headerNames.Count; c++)
                        {
                            var name = headerNames[c - 1];
                            if (cells.TryGetValue(name, out var v))
                                (ws.Cells[row, c] as Excel.Range)!.Value2 = XqlCommon.ValueToString(v);
                        }
                        row++;
                    }
                }
                finally
                {
                    XqlCommon.ReleaseCom(header, ws);
                }
            }
        }

        // 기존 증분 적용기 호출만 남긴 껍데기(가독성용)
        private Task ApplyIncrementalPatches(PullResult pr)
        {
            if (pr?.Patches is { Count: > 0 })
                XqlSheetView.ApplyOnUiThread(pr.Patches);
            return Task.CompletedTask;
        }

        // using Excel = Microsoft.Office.Interop.Excel;
        // using System.Linq;
        // using System.Collections.Generic;

        internal static void InternalApplyCore(Excel.Application app, IEnumerable<RowPatch> patches)
        {
            if (patches == null) return;

            var byTable = patches
                .Where(p => !string.IsNullOrWhiteSpace(p.Table))
                .GroupBy(p => p.Table!, StringComparer.OrdinalIgnoreCase);

            foreach (var grp in byTable)
            {
                Excel.Worksheet? ws = null;
                Excel.Range? header = null;
                Excel.Range? data = null;

                try
                {
                    // ── 1) 테이블→시트 매핑 확보(대소문자 무시, 메타 없으면 즉시 생성)
                    XqlSheet.Meta? smeta;
                    ws = XqlSheetView.FindWorksheetByTable(app, grp.Key, out smeta);
                    if (ws == null)
                    {
                        // 시트가 없으면 생성(부트스트랩 시나리오)
                        ws = (Excel.Worksheet)app.Worksheets.Add();
                        ws.Name = grp.Key.Length > 31 ? grp.Key.Substring(0, 31) : grp.Key;
                        smeta = XqlAddIn.Sheet!.GetOrCreateSheet(ws.Name);
                        XqlSheetView.RegisterTableSheet(grp.Key, ws.Name);
                    }
                    if (smeta == null)
                        smeta = XqlAddIn.Sheet!.GetOrCreateSheet(ws.Name);

                    // ── 2) 헤더/테이블 범위 확보(표가 있으면 HeaderRowRange, 없으면 1행 추정)
                    var lo = XqlSheet.FindListObjectContaining(ws, null);
                    header = lo?.HeaderRowRange ?? XqlSheet.GetHeaderRange(ws);
                    if (header == null)
                    {
                        // 1행 전체를 헤더 후보로
                        header = ws.Range[ws.Cells[1, 1], ws.Cells[1, Math.Max(1, GuessMaxColumnFromUsedRange(ws))]];
                    }

                    // ── 3) 서버 패치에서 등장한 컬럼 수집(+키 보장)
                    var serverCols = new HashSet<string>(StringComparer.Ordinal);
                    foreach (var p in grp)
                    {
                        if (p.Deleted || p.Cells == null) continue;
                        foreach (var k in p.Cells.Keys)
                            if (!string.IsNullOrWhiteSpace(k)) serverCols.Add(k);
                    }
                    var keyName = string.IsNullOrWhiteSpace(smeta.KeyColumn) ? "id" : smeta.KeyColumn!;
                    serverCols.Add(keyName);

                    // ── 4) 현재 헤더 이름 수집
                    var headers = new List<string>(header.Columns.Count);
                    for (int i = 1; i <= header.Columns.Count; i++)
                    {
                        Excel.Range? hc = null;
                        try
                        {
                            hc = (Excel.Range)header.Cells[1, i];
                            var nm = (hc.Value2 as string)?.Trim();
                            headers.Add(string.IsNullOrEmpty(nm)
                                ? XqlCommon.ColumnIndexToLetter(header.Column + i - 1)
                                : nm!);
                        }
                        finally { XqlCommon.ReleaseCom(hc); }
                    }

                    // ── 5) 헤더 자동 생성 필요 여부 판정
                    bool LooksDefaultLetters()
                    {
                        if (headers.Count == 0) return true;
                        for (int i = 0; i < headers.Count; i++)
                        {
                            var expect = XqlCommon.ColumnIndexToLetter(header.Column + i);
                            if (!string.Equals(headers[i], expect, StringComparison.Ordinal)) return false;
                        }
                        return true;
                    }

                    bool needCreateHeader =
                        headers.Count == 0 ||
                        LooksDefaultLetters() ||
                        !headers.Any(h => serverCols.Contains(h));

                    if (needCreateHeader && serverCols.Count > 0)
                    {
                        // 키 우선, 나머지는 알파벳 정렬(안정적 재현)
                        var ordered = new List<string>(serverCols.Count);
                        if (serverCols.Contains(keyName)) ordered.Add(keyName);
                        ordered.AddRange(serverCols.Where(c => !string.Equals(c, keyName, StringComparison.Ordinal))
                                                   .OrderBy(c => c, StringComparer.Ordinal));

                        // 1행에 헤더 텍스트 배치
                        var start = (Excel.Range)header.Cells[1, 1];
                        var end = (Excel.Range)ws.Cells[header.Row, header.Column + ordered.Count - 1];
                        var newHeader = ws.Range[start, end];
                        XqlCommon.ReleaseCom(start, end);

                        var arr = new object[1, ordered.Count];
                        for (int i = 0; i < ordered.Count; i++) arr[0, i] = ordered[i];

                        newHeader.Value2 = arr;

                        // 메타/마커/UI 동기화
                        XqlAddIn.Sheet!.EnsureColumns(ws.Name, ordered);
                        XqlSheet.SetHeaderMarker(ws, newHeader);
                        XqlSheetView.ApplyHeaderUi(ws, newHeader, smeta, withValidation: true);
                        XqlSheetView.InvalidateHeaderCache(ws.Name);
                        XqlSheetView.RegisterTableSheet(grp.Key, ws.Name);

                        // 로컬 변수 교체
                        XqlCommon.ReleaseCom(header);
                        header = newHeader;
                        headers = ordered;
                    }

                    // (헤더를 새로 만들지 않은 경로에서도) 서버 컬럼을 메타에 보장 — 스킵 방지
                    try { XqlAddIn.Sheet!.EnsureColumns(ws.Name, serverCols.ToArray()); } catch { }

                    if (headers.Count == 0) continue;

                    // ── 6) 데이터 영역 계산(표가 있으면 DataBodyRange, 없으면 헤더 아래 ~ 마지막)
                    data = ResolveDataRange(ws, header);

                    // ── 7) 키 절대열 계산
                    int keyIdx1 = XqlSheet.FindKeyColumnIndex(headers, smeta.KeyColumn); // 1-based
                    if (keyIdx1 <= 0) keyIdx1 = 1; // 폴백: 1열
                    int keyAbsCol = header.Column + keyIdx1 - 1;

                    // ── 8) 적용(성능: 화면/계산/이벤트 잠시 off)
                    var prevCalc = app.Calculation;
                    var prevScreen = app.ScreenUpdating;
                    var prevEvt = app.EnableEvents;
                    try
                    {
                        app.Calculation = Excel.XlCalculation.xlCalculationManual;
                        app.ScreenUpdating = false;
                        app.EnableEvents = false;

                        foreach (var p in grp)
                        {
                            if (p.Deleted)
                            {
                                // 삭제: 키로 행 찾고 삭제
                                var rowIx = FindRowByKey(ws, data, keyAbsCol, p.RowKey?.ToString() ?? string.Empty);
                                if (rowIx > 0)
                                {
                                    var delRng = (Excel.Range)ws.Rows[rowIx];
                                    delRng.Delete(Excel.XlDeleteShiftDirection.xlShiftUp);
                                    XqlCommon.ReleaseCom(delRng);
                                }
                                continue;
                            }

                            // 업서트: 키행 찾기 또는 append
                            var r = FindRowByKey(ws, data, keyAbsCol, p.RowKey?.ToString() ?? string.Empty);
                            if (r <= 0)
                            {
                                // 맨 아래 Append
                                r = Math.Max(data.Row, LastUsedRow(ws) + 1);
                                var keyCell = (Excel.Range)ws.Cells[r, keyAbsCol];
                                keyCell.Value2 = p.RowKey;
                                XqlCommon.ReleaseCom(keyCell);
                            }

                            if (p.Cells != null)
                            {
                                foreach (var kv in p.Cells)
                                {
                                    var colName = kv.Key;
                                    var val = kv.Value;

                                    var cIdx = headers.FindIndex(h => string.Equals(h, colName, StringComparison.Ordinal));
                                    if (cIdx < 0)
                                    {
                                        // 새 컬럼이 패치에 등장 → 우측 확장
                                        cIdx = headers.Count;
                                        headers.Add(colName);

                                        // 헤더 확장
                                        var hCell = (Excel.Range)header.Cells[1, cIdx + 1];
                                        hCell.Value2 = colName;
                                        XqlCommon.ReleaseCom(hCell);

                                        // 메타/검증 갱신
                                        try { XqlAddIn.Sheet!.EnsureColumns(ws.Name, headers); } catch { }
                                        XqlSheetView.ApplyHeaderUi(ws, header, smeta, withValidation: true);
                                        XqlSheetView.InvalidateHeaderCache(ws.Name);
                                    }

                                    var absCol = header.Column + cIdx;
                                    Excel.Range? cell = null;
                                    try
                                    {
                                        cell = (Excel.Range)ws.Cells[r, absCol];
                                        cell.Value2 = CoerceForExcel(val);
                                        XqlSheetView.MarkTouchedCell(cell); // 선택: 시각 피드백
                                    }
                                    finally { XqlCommon.ReleaseCom(cell); }
                                }
                            }
                        }
                    }
                    finally
                    {
                        app.EnableEvents = prevEvt;
                        app.ScreenUpdating = prevScreen;
                        app.Calculation = prevCalc;
                    }
                }
                catch
                {
                    // swallow and continue next table
                }
                finally
                {
                    XqlCommon.ReleaseCom(data, header, ws);
                }
            }

            // ────────────────────── 지역 헬퍼들 ──────────────────────

            static int GuessMaxColumnFromUsedRange(Excel.Worksheet ws)
            {
                try { return ws.UsedRange?.Columns?.Count ?? 1; } catch { return 1; }
            }

            static Excel.Range ResolveDataRange(Excel.Worksheet ws, Excel.Range header)
            {
                var lo = XqlSheet.FindListObjectContaining(ws, header);
                if (lo?.DataBodyRange != null) return lo.DataBodyRange;

                var first = (Excel.Range)header.Offset[1, 0];
                var last = ws.Cells[ws.Rows.Count, header.Column + header.Columns.Count - 1];
                var data = ws.Range[first, last];
                XqlCommon.ReleaseCom(first, last, lo);
                return data;
            }

            static int LastUsedRow(Excel.Worksheet ws)
            {
                try
                {
                    var ur = ws.UsedRange;
                    var r = ur.Row + ur.Rows.Count - 1;
                    XqlCommon.ReleaseCom(ur);
                    return Math.Max(1, r);
                }
                catch { return 1; }
            }

            static int FindRowByKey(Excel.Worksheet ws, Excel.Range data, int keyAbsCol, string rowKey)
            {
                var start = data.Row;
                var end = Math.Max(start, LastUsedRow(ws));
                for (int r = start; r <= end; r++)
                {
                    Excel.Range? c = null;
                    try
                    {
                        c = (Excel.Range)ws.Cells[r, keyAbsCol];
                        var v = (c.Value2 as string) ?? c.Value2?.ToString();
                        if (string.Equals(v, rowKey, StringComparison.Ordinal)) return r;
                    }
                    finally { XqlCommon.ReleaseCom(c); }
                }
                return -1;
            }

            static object? CoerceForExcel(object? v)
            {
                if (v is null) return null;
                return v switch
                {
                    bool b => b,
                    double d => d,
                    float f => (double)f,
                    decimal m => (double)m,
                    long l => (double)l,
                    int i => (double)i,
                    DateTime dt => dt.ToOADate(),
                    _ => v.ToString()
                };
            }
        }

        // ⬇️ [ADD] 초기 부트스트랩 판단: 메타 마커 없음, 거의 비어있음, 혹은 A/B/C… 폴백 헤더
        private bool IsInitialBootstrapContext()
        {
            try
            {
                var app = (Excel.Application)ExcelDnaUtil.Application;
                if (app.ActiveSheet is not Excel.Worksheet ws) return false;
                if (XqlSheet.TryGetHeaderMarker(ws, out var hdr)) { XqlCommon.ReleaseCom(hdr); return false; }

                Excel.Range? used = null;
                try
                {
                    used = ws.UsedRange;
                    if ((long)(used?.CountLarge ?? 0L) <= 1) return true;
                }
                finally { XqlCommon.ReleaseCom(used); }

                var header = XqlSheet.GetHeaderRange(ws);
                try
                {
                    int cols = header.Columns.Count;
                    if (cols <= 0) return true;
                    for (int i = 1; i <= cols; i++)
                    {
                        Excel.Range? c = null;
                        try
                        {
                            c = (Excel.Range)header.Cells[1, i];
                            var name = (c.Value2 as string)?.Trim() ?? "";
                            var exp = XqlCommon.ColumnIndexToLetter(header.Column + i - 1);
                            if (!string.Equals(name, exp, StringComparison.Ordinal))
                                return false;
                        }
                        finally { XqlCommon.ReleaseCom(c); }
                    }
                    return true; // 전부 폴백
                }
                finally { XqlCommon.ReleaseCom(header); }
            }
            catch { return false; }
        }

        // ⬇️ [ADD] 현재 시트의 테이블명 계산(메타 없으면 시트명 사용)
        private (string table, string wsName) ResolveActiveTable()
        {
            try
            {
                var app = (Excel.Application)ExcelDnaUtil.Application;
                if (app.ActiveSheet is not Excel.Worksheet ws) return ("", "");
                var sm = XqlAddIn.Sheet?.GetOrCreateSheet(ws.Name);
                var table = string.IsNullOrWhiteSpace(sm?.TableName) ? ws.Name : sm!.TableName!;
                return (table, ws.Name);
            }
            catch { return ("", ""); }
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
                {
#pragma warning disable CS4014 // 이 호출을 대기하지 않으므로 호출이 완료되기 전에 현재 메서드가 계속 실행됩니다.
                    PullSince(before); // 갭 보정
#pragma warning restore CS4014 // 이 호출을 대기하지 않으므로 호출이 완료되기 전에 현재 메서드가 계속 실행됩니다.
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
