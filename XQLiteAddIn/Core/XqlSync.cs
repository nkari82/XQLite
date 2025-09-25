// XqlSync.cs  (ExcelPatchApplier 포함 버전)
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Threading.Tasks;
using ExcelDna.Integration;
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

        // ⬇️ 엑셀 반영기
        private readonly ExcelPatchApplier _applier;

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
            try { _backend.Dispose(); } catch { }
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        public void EnqueueCellEdit(string table, object rowKey, string column, object? value)
        {
            if (_disposed) return;
            _outbox.Enqueue(new EditCell(table, rowKey, column, value));
        }

        public void FlushUpsertsNow() => SafeFlushUpserts();

        public Task<PullResult?> PullSince() => SafePull(MaxRowVersion);

        private void SafeFlushUpserts()
        {
            if (!_started || _disposed) return;

            lock (_flushGate)
            {
                try
                {
                    // 비동기 코어 실행 (락 밖에서 await되도록 async void 유지)
                    FlushUpsertsCore();
                }
                catch (Exception ex)
                {
                    _conflicts.Enqueue(Conflict.System("flush", ex.Message));
                }
            }
        }

        private async Task<PullResult?> SafePull(long? sinceOverride = null)
        {
            if (!_started || _disposed)
                return null;

            Task<PullResult>? result;

            lock (_pullGate)
            {
                try
                {
                    result = PullCore(sinceOverride ?? MaxRowVersion);
                }
                catch (Exception ex)
                {
                    _conflicts.Enqueue(Conflict.System("pull", ex.Message));
                    return null;
                }
            }

            return await result.ConfigureAwait(false);
        }

        private async void FlushUpsertsCore()
        {
            try
            {
                if (_outbox.IsEmpty) return;

                var batch = DrainDedupCells(_outbox, 512);
                if (batch.Count == 0) return;

                var resp = await _backend.UpsertCells(batch).ConfigureAwait(false);

                if (resp.Errors?.Count > 0)
                    foreach (var e in resp.Errors)
                        _conflicts.Enqueue(Conflict.System("upsert", e));

                if (resp.MaxRowVersion > 0)
                    XqlCommon.InterlockedMax(ref _maxRowVersion, resp.MaxRowVersion);

                if (resp.Conflicts is { Count: > 0 })
                    foreach (var c in resp.Conflicts)
                        _conflicts.Enqueue(c);
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
                if (ev.MaxRowVersion > 0)
                    XqlCommon.InterlockedMax(ref _maxRowVersion, ev.MaxRowVersion);

                // ⬇️ 푸시 패치 즉시 적용 (UI 스레드)
                if (ev.Patches is { Count: > 0 })
                    _applier.ApplyOnUiThread(ev.Patches);

                // 안전성 위해 한 번 더 Pull (fire-and-forget)
                var _ = SafePull();
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
            return [.. map.Values];
        }

        private readonly record struct CellKey(string Table, object RowKey, string Column);

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
                    try
                    {
                        ws = FindWorksheetByTable(app, grp.Key, out var smeta);
                        if (ws == null || smeta == null) continue;

                        // 헤더/컬럼 맵
                        var (header, headers) = XqlSheet.GetHeaderAndNames(ws);
                        if (headers.Count == 0) continue;

                        // 키 컬럼 인덱스 결정 (메타 우선)
                        int keyCol = XqlSheet.FindKeyColumnIndex(headers, smeta.KeyColumn);
                        int firstDataRow = header.Row + 1;

                        foreach (var patch in grp)
                        {
                            try
                            {
                                int? row = XqlSheet.FindRowByKey(ws, firstDataRow, keyCol, patch.RowKey);
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
                    finally { XqlCommon.ReleaseCom(ws); }
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
                            if (_sheet.TryGetSheet(name, out var m))
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
                        finally { XqlCommon.ReleaseCom(w); }
                    }
                }
                catch { }

                return null;
            }

            private static int AppendNewRow(Excel.Worksheet ws, int firstDataRow)
            {
                int last = firstDataRow;
                try
                {
                    var used = ws.UsedRange;
                    last = used.Row + used.Rows.Count - 1;
                    XqlCommon.ReleaseCom(used);
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
