// XqlSync.cs  (ExcelPatchApplier 포함 버전)
using ExcelDna.Integration;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace XQLite.AddIn
{
    internal sealed class XqlSync : IDisposable
    {
        private readonly int _pushIntervalMs;
        private readonly int _pullIntervalMs;

        private readonly IXqlBackend _backend;
        private readonly XqlSheet _sheet;
        private readonly ConcurrentQueue<EditCell> _outbox = new();
        private readonly SemaphoreSlim _pushSem = new(1, 1);
        private readonly SemaphoreSlim _pullSem = new(1, 1);
        private volatile bool _pullAgain;

        private long _maxRowVersion;
        public long MaxRowVersion => Interlocked.Read(ref _maxRowVersion);

        private readonly Timer _pushTimer;
        private readonly Timer _pullTimer;

        private volatile bool _started;
        private volatile bool _disposed;

        // ⬇️ 엑셀 반영기
        private readonly ExcelPatchApplier _applier;

        private const int UPSERT_CHUNK = 512;   // 1회 전송 셀 수
        private const int UPSERT_SLICE_MS = 250; // 한번에 잡는 시간

        private readonly ConcurrentQueue<Conflict> _conflicts = new();

        private readonly ConcurrentDictionary<string, string?> _lastPushed = new(StringComparer.Ordinal);
        private static string Key(EditCell e) => $"{e.Table}\n{XqlCommon.ValueToString(e.RowKey)}\n{e.Column}";

        // Excel 값 정규화: null/빈문자/숫자/불린/date를 문자열로 안정 변환
        private static string? Canon(object? v)
        {
            if (v is null) return null;
            if (v is bool b) return b ? "1" : "0";
            if (v is double d) return d.ToString("R", System.Globalization.CultureInfo.InvariantCulture); // Value2 숫자
            if (v is DateTime dt) return dt.ToUniversalTime().Ticks.ToString(); // 필요 시 epoch ms로 변경
            var s = v.ToString();
            return string.IsNullOrWhiteSpace(s) ? null : s;
        }

        public void EnqueueIfChanged(string table, string rowKey, string column, object? value)
        {
            var e = new EditCell { Table = table, RowKey = rowKey, Column = column, Value = value };
            var k = Key(e);
            var norm = Canon(value);
            if (_lastPushed.TryGetValue(k, out var prev) && prev == norm)
                return; // 동일값 → 전송 생략
            _outbox.Enqueue(e);
        }

        public bool TryDequeueConflict(out Conflict c) => _conflicts.TryDequeue(out c);


        public XqlSync(IXqlBackend backend, XqlSheet sheet, int pushIntervalMs = 2000, int pullIntervalMs = 10000)
        {
            _sheet = sheet ?? throw new ArgumentNullException(nameof(sheet));
            _pushIntervalMs = Math.Max(250, pushIntervalMs);
            _pullIntervalMs = Math.Max(1000, pullIntervalMs);

            _backend = backend ?? throw new ArgumentNullException(nameof(backend));
            _applier = new ExcelPatchApplier(_sheet);

            _pushTimer = new Timer(_ => SafeFlushUpserts(), null, Timeout.Infinite, Timeout.Infinite);
            _pullTimer = new Timer(_ => _ = SafePull(), null, Timeout.Infinite, Timeout.Infinite);
        }

        public void Start()
        {
            if (_disposed || _started) return;
            _started = true;

            _pushTimer.Change(_pushIntervalMs, _pushIntervalMs);
            _pullTimer.Change(_pullIntervalMs, _pullIntervalMs);

            // ✅ 구독 시작은 동기 메서드 사용
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
        }

        public void FlushUpsertsNow() => _ = FlushUpsertsNow(false);

        public async Task PullSince(long? sinceOverride = null)
        {
            if (!await _pullSem.WaitAsync(0).ConfigureAwait(false))
            {
                _pullAgain = true; // 이미 실행 중 → 한 번 더 해달라 표시
                return;
            }
            try
            {
                do
                {
                    _pullAgain = false;
                    var since = sinceOverride ?? MaxRowVersion;
                    var pr = await _backend.PullRows(since).ConfigureAwait(false);

                    if (pr.Patches is { Count: > 0 })
                        _applier.ApplyOnUiThread(pr.Patches);

                    if (pr.MaxRowVersion > 0)
                        XqlCommon.InterlockedMax(ref _maxRowVersion, pr.MaxRowVersion);
                }
                while (_pullAgain); // 실행 중 요청이 있었으면 즉시 최신 기준으로 1회 더
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
                    catch (Exception ex) { _conflicts.Enqueue(Conflict.System("upsert.core", ex.Message)); }
                    finally { _pushSem.Release(); }
                });
            }
            catch (Exception ex)
            {
                _conflicts.Enqueue(Conflict.System("flush", ex.Message));
            }
        }

        private async Task<PullResult?> SafePull(long? sinceOverride = null)
        {
            if (!_started || _disposed) return null;
            if (!_pullSem.Wait(0)) return null;
            var task = PullCore(sinceOverride ?? MaxRowVersion);
            try { return await task.ConfigureAwait(false); }
            catch (Exception ex) { _conflicts.Enqueue(Conflict.System("pull", ex.Message)); return null; }
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

                    var resp = await _backend.UpsertCells(batch).ConfigureAwait(false);

                    // ✅ 성공 반영값을 스냅샷에 기록(이후 동일값 재전송 방지)
                    foreach (var e in batch)
                        _lastPushed[Key(e)] = Canon(e.Value);

                    if (resp.Errors is { Count: > 0 })
                        foreach (var e in resp.Errors)
                            _conflicts.Enqueue(Conflict.System("upsert", e));

                    if (resp.MaxRowVersion > 0)
                        XqlCommon.InterlockedMax(ref _maxRowVersion, resp.MaxRowVersion);

                    if (resp.Conflicts is { Count: > 0 })
                        foreach (var c in resp.Conflicts) _conflicts.Enqueue(c);
                }
                while (XqlCommon.Monotonic.NowMs() < deadline && !_outbox.IsEmpty);
            }
            catch (Exception ex)
            {
                _conflicts.Enqueue(Conflict.System("upsert.core", ex.Message));
            }
        }
        private async Task<PullResult> PullCore(long sinceVersion)
        {
            var resp = await _backend.PullRows(sinceVersion).ConfigureAwait(false);

            if (resp.MaxRowVersion > 0)
                XqlCommon.InterlockedMax(ref _maxRowVersion, resp.MaxRowVersion);

            // ⬇️ 서버 패치를 엑셀에 적용 (UI 스레드 매크로 큐로 안전하게)
            if (resp.Patches is { Count: > 0 })
                _applier.ApplyOnUiThread(resp.Patches);

            return resp;
        }

        private void OnServerEvent(ServerEvent ev)
        {
            try
            {
                var before = MaxRowVersion;

                if (ev.Patches is { Count: > 0 })
                    _applier.ApplyOnUiThread(ev.Patches);

                if (ev.MaxRowVersion > 0)
                    XqlCommon.InterlockedMax(ref _maxRowVersion, ev.MaxRowVersion);

                if (ev.MaxRowVersion > before + 1)
                    _ = PullSince(before); // 갭 보정
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

            // 튜플 comparer 캐스팅 대신 안정적인 문자열 키 사용
            static string K(EditCell e) => $"{e.Table}\n{XqlCommon.ValueToString(e.RowKey)}\n{e.Column}";

            var map = new Dictionary<string, EditCell>(temp.Count, StringComparer.Ordinal);
            foreach (var e in temp) map[K(e)] = e; // 마지막 값 우선
            return map.Values.ToList();
        }

        // ========== ⬇️ 엑셀 반영기: 서버 패치 → 시트 적용 (UI 스레드에서 실행) ==========

        private sealed class ExcelPatchApplier
        {
            private readonly XqlSheet _sheet;
            public ExcelPatchApplier(XqlSheet sheet) => _sheet = sheet;

            public void ApplyOnUiThread(List<RowPatch> patches)
            {
                if (patches == null || patches.Count == 0) return;
                ExcelAsyncUtil.QueueAsMacro(() =>
                {
                    try { ApplyNow(patches); } catch { }
                });
            }

            private void ApplyNow(List<RowPatch> patches)
            {
                var app = (Excel.Application)ExcelDnaUtil.Application;

                // 테이블별 그룹
                foreach (var grp in patches.GroupBy(p => p.Table, StringComparer.Ordinal))
                {
                    Excel.Worksheet? ws = null;

                    ws = FindWorksheetByTable(app, grp.Key, out var smeta);
                    if (ws == null || smeta == null) continue;

                    // 헤더/컬럼 맵
                    Excel.Range? header = null; Excel.ListObject? lo = null;
                    try
                    {
                        lo = XqlSheet.FindListObjectByTable(ws, grp.Key);
                        header = lo?.HeaderRowRange ?? XqlSheet.GetHeaderRange(ws);
                        var headers = new List<string>(header.Columns.Count);
                        for (int i = 1; i <= header.Columns.Count; i++)
                        {
                            Excel.Range? hc = null; try
                            {
                                hc = (Excel.Range)header.Cells[1, i];
                                var nm = (hc.Value2 as string)?.Trim();
                                headers.Add(string.IsNullOrEmpty(nm) ? XqlCommon.ColumnIndexToLetter(header.Column + i - 1) : nm!);
                            }
                            finally { XqlCommon.ReleaseCom(hc); }
                        }
                        if (headers.Count == 0) continue;

                        // 키 컬럼 인덱스 결정 (메타 우선)
                        int keyIdx1 = XqlSheet.FindKeyColumnIndex(headers, smeta.KeyColumn); // 1-based
                        int keyAbsCol = header.Column + keyIdx1 - 1;                         // 절대열
                        int firstDataRow = header.Row + 1;

                        foreach (var patch in grp)
                        {
                            try
                            {
                                int? row = XqlSheet.FindRowByKey(ws, firstDataRow, keyAbsCol, patch.RowKey);
                                if (patch.Deleted)
                                {
                                    if (row.HasValue) SafeDeleteRow(ws, row.Value);
                                    continue;
                                }
                                if (!row.HasValue) row = AppendNewRow(ws, firstDataRow, lo);

                                ApplyCells(ws, row!.Value, header, headers, smeta, patch.Cells);
                            }
                            catch { /* per-row safe */ }
                        }


                    }
                    finally { XqlCommon.ReleaseCom(lo); XqlCommon.ReleaseCom(header); XqlCommon.ReleaseCom(ws); }
                }
            }

            // === 메타 기반: 테이블명 → 워크시트 찾기 ===
            private Excel.Worksheet? FindWorksheetByTable(Excel.Application app, string table, out SheetMeta? smeta)
            {
                smeta = null;

                Excel.Worksheet? match = null;
                // 1) 시트명 == 테이블명인 경우
                try
                {
                    foreach (Excel.Worksheet w in app.Worksheets)
                    {
                        try
                        {
                            string name = w.Name;
                            // 메타가 등록된 시트만 대상
                            if (_sheet.TryGetSheet(name, out var m))
                            {
                                if (string.Equals(m.TableName ?? name, table, StringComparison.Ordinal))
                                {
                                    smeta = m;
                                    match = w;
                                    break;
                                }
                                // 시트명 자체가 테이블명인 케이스도 통과
                                if (string.Equals(name, table, StringComparison.Ordinal) && smeta == null)
                                {
                                    smeta = m;
                                    match = w;
                                    break;
                                }
                            }
                        }
                        finally
                        {
                            if (!object.ReferenceEquals(match, w)) XqlCommon.ReleaseCom(w);
                        }
                    }
                }
                catch { }

                return match;
            }

            private static int AppendNewRow(Excel.Worksheet ws, int firstDataRow, Excel.ListObject? lo)
            {
                // 1) 표가 있으면 표에 행 추가(범위 자동 확장)
                if (lo != null)
                {
                    try
                    {
                        var lr = lo.ListRows.Add();    // 테이블 마지막에 1행 추가
                        XqlCommon.ReleaseCom(lr);
                        var body = lo.DataBodyRange;   // 새로고침된 바디 참조
                        if (body != null)
                        {
                            int row = body.Row + body.Rows.Count - 1;
                            XqlCommon.ReleaseCom(body);
                            return row;
                        }
                    }
                    catch { /* 실패 시 폴백 */ }
                }
                // 2) 폴백: UsedRange 기반으로 시트에 직접 추가
                int last = firstDataRow;
                try { var used = ws.UsedRange; last = used.Row + used.Rows.Count - 1; XqlCommon.ReleaseCom(used); }
                catch { }
                return Math.Max(firstDataRow, last + 1);
            }

            // ✅ 메타 컬럼에 정의된 컬럼만 적용
            private static void ApplyCells(Excel.Worksheet ws, int row, Excel.Range header, List<string> headers, SheetMeta meta, Dictionary<string, object?> cells)
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
                        rg = (Excel.Range)ws.Cells[row, header.Column + c]; // ✅ 헤더 절대열 기준
                        if (val == null) { rg.Value2 = null; continue; }

                        switch (val)
                        {
                            case bool b: rg.Value2 = b; break;
                            case long l: rg.Value2 = (double)l; break;
                            case int i: rg.Value2 = (double)i; break;
                            case double d: rg.Value2 = d; break;
                            case float f: rg.Value2 = (double)f; break;
                            case decimal m: rg.Value2 = (double)m; break;
                            case DateTime dt: rg.Value2 = dt.ToOADate(); break;
                            default: rg.Value2 = Convert.ToString(val, System.Globalization.CultureInfo.InvariantCulture); break;
                        }

                        // 서버 패치로 변경된 셀 하이라이트
                        XqlSheetView.MarkTouchedCell(rg);
                    }
                    catch (Exception ex)
                    {
                        XqlLog.Error($"패치 적용 실패: {ex.Message}", ws.Name,
                        rg?.Address[false, false] ?? "");
                    }
                    finally { XqlCommon.ReleaseCom(rg); }
                }
            }

            private static void SafeDeleteRow(Excel.Worksheet ws, int row)
            {
                try { var rg = (Excel.Range)ws.Rows[row]; rg.Delete(); XqlCommon.ReleaseCom(rg); }
                catch { }
            }
        }
    }
}
