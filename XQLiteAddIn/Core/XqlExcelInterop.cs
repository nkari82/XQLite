// XqlExcelInterop.cs
using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace XQLite.AddIn
{
    /// <summary>
    /// Excel 개체모델과 Add-in 내부 모듈(XqlSync/XqlCollab/XqlMetaRegistry)을 연결한다.
    /// - 리본/메뉴 → 명령 라우팅(Cmd_*)
    /// - Excel 이벤트(시트 변경/선택 변경/통합문서 열기·닫기) 핸들링
    /// - 헤더 범위 탐색, 타입 툴팁/주석 표시
    /// - 셀 검증(정적 로직은 XqlMetaRegistry), 결과 시각화(주석)
    /// - Presence/락 하트비트 전송(선택 변경 시)
    /// - 셀 편집 → 2초 디바운스 업서트 큐 적재(XqlSync)
    /// </summary>
    internal sealed class XqlExcelInterop(Excel.Application app, XqlSync sync, XqlCollab collab, XqlSheet sheet, XqlBackup backup) : IDisposable
    {
        private readonly Excel.Application _app = app ?? throw new ArgumentNullException(nameof(app));
        private readonly XqlSync _sync = sync ?? throw new ArgumentNullException(nameof(sync));
        private readonly XqlCollab _collab = collab ?? throw new ArgumentNullException(nameof(collab));
        private readonly XqlSheet _sheet = sheet ?? throw new ArgumentNullException(nameof(sheet));
        private readonly XqlBackup _backup = backup ?? throw new ArgumentNullException(nameof(backup));

        private bool _started;

        // ========= 수명 주기 =========

        public void Start()
        {
            if (_started) return;
            _started = true;

            try
            {
                var app = ExcelDnaUtil.Application as Excel.Application;
                var wb = app?.ActiveWorkbook;
                if (wb != null)
                {
                    var full = wb.FullName; // c:\path\file.xlsx
                    _sync.InitPersistentState(full, XqlConfig.Project);
                }
            }
            catch { /* 무시 */ }

            _app.SheetChange += App_SheetChange;
            _app.SheetSelectionChange += App_SheetSelectionChange;
            _app.WorkbookOpen += App_WorkbookOpen;
            _app.WorkbookBeforeClose += App_WorkbookBeforeClose;
        }

        public void Stop()
        {
            if (!_started) return;
            _started = false;

            _app.SheetChange -= App_SheetChange;
            _app.SheetSelectionChange -= App_SheetSelectionChange;
            _app.WorkbookOpen -= App_WorkbookOpen;
            _app.WorkbookBeforeClose -= App_WorkbookBeforeClose;
        }

        public void Dispose()
        {
            Stop();
        }

        // ========= 명령(리본/메뉴) =========
        public async void Cmd_CommitSync()
        {
            try
            {
                await EnsureActiveSheetSchema();
                await _sync.FlushUpsertsNow(force: true);  // 변경된 셀만 즉시 업서트
                                                           // Pull은 필요 시(버튼/갭 감지/주기)만
            }
            catch (Exception ex)
            {
                XqlLog.Warn("CommitSmart failed: " + ex.Message);
            }
        }

        // XqlExcelInterop.cs

        public async void Cmd_PullOnly()
        {
            try
            {
                var app = ExcelDnaUtil.Application as Excel.Application;
                if (app == null) { await _sync.PullSince(); return; }

                var ws = app.ActiveSheet as Excel.Worksheet;
                if (ws == null) { await _sync.PullSince(); return; }

                // ✅ 부트스트랩 필요 판단: 마커 없음 OR 유효 데이터 거의 없음 OR 1행이 A/B/C… 폴백 헤더
                bool needsBootstrap = XqlSheet.NeedsBootstrap(ws);

                // 강제 Full Pull(since=0) 또는 증분 Pull
                await _sync.PullSince(needsBootstrap ? 0 : (long?)null);
            }
            catch (Exception ex)
            {
                XqlLog.Warn("PullOnly failed: " + ex.Message);
            }
        }

        public void Cmd_RecoverFromExcel()
        {
            // 원클릭 복구: 엑셀 파일을 원본으로 DB 재생성
            var _ = _backup.RecoverFromExcel();
        }

        // ========= Excel 이벤트 =========

        private void App_WorkbookOpen(Excel.Workbook wb)
        {
            try
            {
                if (wb != null)
                    _sync.InitPersistentState(wb.FullName, XqlConfig.Project);
            }
            catch { /* ignore */ }
            finally { XqlCommon.ReleaseCom(wb); }
        }


        private void App_WorkbookBeforeClose(Excel.Workbook wb, ref bool Cancel)
        {
            // 락 해제, 프레즌스 정리, 타이머 정지(중복 안전)
            var _ = _collab.ReleaseByMe();
            try { _sync.Stop(); } catch { }
            XqlCommon.ReleaseCom(wb);
        }

        /// <summary>시트에서 헤더와 데이터 범위를 일관되게 구한다.</summary>
        private static (Excel.Range? header, Excel.Range? data, Excel.ListObject? lo) ResolveHeaderAndData(Excel.Worksheet sh)
        {
            Excel.Range? header = XqlSheetView.ResolveHeader(sh, null, XqlAddIn.Sheet!)
                          ?? (XqlSheet.TryGetHeaderMarker(sh, out var mk) ? mk : XqlSheet.GetHeaderRange(sh));
            if (header == null) return (null, null, null);
            var lo = XqlSheet.FindListObjectContaining(sh, header);
            if (lo?.DataBodyRange != null) return (header, lo.DataBodyRange, lo);
            var first = (Excel.Range)header.Offset[1, 0];
            var last = sh.Cells[sh.Rows.Count, header.Column + header.Columns.Count - 1];
            var data = sh.Range[first, last];
            XqlCommon.ReleaseCom(first, last);
            return (header, data, lo);
        }

        // 변경 이벤트에서 호출
        private void App_SheetChange(object Sh, Excel.Range target)
        {
            Excel.Worksheet? sh = null;
            Excel.Range? header = null;
            Excel.Range? data = null;
            Excel.Range? intersect = null;
            Excel.ListObject? lo = null;

            try
            {
                sh = Sh as Excel.Worksheet;
                if (sh == null) return;

                var sm = _sheet.GetOrCreateSheet(sh.Name);
                (header, data, lo) = ResolveHeaderAndData(sh);
                if (header == null || data == null) return;

                // 헤더 편집이면 캐시 무효화만
                var hitHeader = sh.Application.Intersect(target, header) as Excel.Range;
                if (hitHeader != null)
                {
                    XqlSheetView.InvalidateHeaderCache(sh.Name);
                    XqlCommon.ReleaseCom(hitHeader);
                    return;
                }

                intersect = sh.Application.Intersect(target, data) as Excel.Range;
                if (intersect == null) return;

                var table = string.IsNullOrWhiteSpace(sm.TableName) ? sh.Name : sm.TableName!;
                var keyColName = string.IsNullOrWhiteSpace(sm.KeyColumn) ? "id" : sm.KeyColumn!;

                var areas = intersect.Areas;
                try
                {
                    for (int ai = 1; ai <= areas.Count; ai++)
                    {
                        Range? area = null;
                        try
                        {
                            area = (Excel.Range)areas[ai];
                            foreach (Excel.Range cell in area.Cells)
                            {
                                Excel.Range? hdrCell = null;
                                Excel.Range? keyCell = null;
                                try
                                {
                                    int hdrIdx = cell.Column - header.Column + 1;
                                    if (hdrIdx < 1 || hdrIdx > header.Columns.Count) continue;

                                    hdrCell = (Excel.Range)header.Cells[1, hdrIdx];
                                    var colName = (hdrCell.Value2 as string)?.Trim();
                                    if (string.IsNullOrWhiteSpace(colName))
                                        colName = XqlCommon.ColumnIndexToLetter(cell.Column);

                                    int keyAbsCol = XqlSheet.FindKeyColumnAbsolute(header, sm.KeyColumn);
                                    keyCell = sh.Cells[cell.Row, keyAbsCol] as Excel.Range;

                                    var rowKeyObj = keyCell?.Value2;
                                    string? rowKey = rowKeyObj?.ToString();

#if false
                                    if (string.IsNullOrWhiteSpace(rowKey))
                                    {
                                        if (!string.Equals(colName, keyColName, StringComparison.OrdinalIgnoreCase))
                                            continue;

                                        rowKey = Guid.NewGuid().ToString("N").Substring(0, 12);
                                        if (keyCell != null) keyCell.Value2 = rowKey;
                                    }
#else
                                    // ✅ 키가 비어 있으면, 수정한 컬럼이 무엇이든 키를 자동 생성해 채운다.
                                    if (string.IsNullOrWhiteSpace(rowKey))
                                    {
                                        rowKey = Guid.NewGuid().ToString("N").Substring(0, 12);
                                        if (keyCell != null) keyCell.Value2 = rowKey;
                                    }
#endif

                                    object? value = cell.Value2;
                                    _sync.EnqueueIfChanged(table, rowKey!, colName!, value);
                                }
                                finally
                                {
                                    XqlCommon.ReleaseCom(keyCell, hdrCell, cell);
                                }
                            }
                        }
                        finally { XqlCommon.ReleaseCom(area); }
                    }
                }
                finally { XqlCommon.ReleaseCom(areas); }
            }
            catch (Exception ex)
            {
                XqlLog.Warn("OnWorksheetChange: " + ex.Message);
            }
            finally { XqlCommon.ReleaseCom(intersect, data, lo, header, sh); }
        }

        private void App_SheetSelectionChange(object Sh, Excel.Range Target)
        {
            try
            {
                var ws = Sh as Excel.Worksheet; if (ws == null) return;
                string sheet = ws.Name;
                string cell = Target?.Address[false, false] ?? "";

                XqlAddIn.Collab?.SelectionChanged(sheet, cell);
            }
            catch { /* non-fatal */ }
            finally { XqlCommon.ReleaseCom(Target); }
        }

        /// <summary>
        /// 활성 시트의 헤더/메타를 읽어 서버에 테이블 없으면 생성, 누락 컬럼을 추가.
        /// </summary>
        private async Task EnsureActiveSheetSchema()
        {
            if (XqlAddIn.Backend is not IXqlBackend be) return;

            var app = ExcelDnaUtil.Application as Excel.Application;
            if (app == null) return;

            Excel.Worksheet? ws = null; Excel.Range? header = null;
            try
            {
                ws = app.ActiveSheet as Excel.Worksheet;
                if (ws == null) return;

                var sm = _sheet.GetOrCreateSheet(ws.Name);

                // ① 헤더/컬럼명
                var (hdr, names) = XqlSheet.GetHeaderAndNames(ws);
                header = hdr;
                if (header == null || names is not { Count: > 0 }) return;

                _sheet.EnsureColumns(ws.Name, names);

                var table = string.IsNullOrWhiteSpace(sm.TableName) ? ws.Name : sm.TableName!;
                var key = string.IsNullOrWhiteSpace(sm.KeyColumn) ? "id" : sm.KeyColumn!;

                // ② 테이블 보장
                await be.TryCreateTable(table, key);

                // ✅ 새 테이블이 생겼거나 헤더 기본구성이 바뀌었을 수 있음 → 캐시 무효화
                XqlSheetView.InvalidateHeaderCache(ws.Name);

                // ③ 서버 컬럼 조회 → '없을 때만' 추가
                var serverCols = await be.GetTableColumns(table);
                var serverSet = new HashSet<string>(serverCols.Select(c => c.name), StringComparer.OrdinalIgnoreCase);
                var addTargets = sm.Columns.Keys.Where(k => !serverSet.Contains(k)).ToList();

                if (addTargets.Count > 0)
                {
                    var defs = addTargets.Select(name =>
                    {
                        var ct = sm.Columns[name];
                        return new ColumnDef
                        {
                            Name = name,
                            Kind = ct.Kind switch
                            {
                                XqlSheet.ColumnKind.Int => "integer",
                                XqlSheet.ColumnKind.Real => "real",
                                XqlSheet.ColumnKind.Bool => "bool",     // ✅ 서버 규격
                                XqlSheet.ColumnKind.Date => "integer",  // (epoch ms를 int로 저장)
                                XqlSheet.ColumnKind.Json => "json",
                                _ => "text"
                            },
                            NotNull = !ct.Nullable,
                            Check = null
                        };
                    });
                    await be.TryAddColumns(table, defs);

                    // ✅ 헤더 컬럼이 늘어남 → 캐시 무효화
                    XqlSheetView.InvalidateHeaderCache(ws.Name);
                }

                // ④ (옵션) 헤더에 없는 서버 컬럼 DROP
                if (XqlConfig.DropColumnsOnCommit)
                {
                    var metaCols = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { key, "row_version", "updated_at", "deleted" };
                    var headerSet = new HashSet<string>(names, StringComparer.OrdinalIgnoreCase);

                    var drop = serverCols
                        .Where(c => !c.pk && !metaCols.Contains(c.name))
                        .Select(c => c.name)
                        .Where(n => !headerSet.Contains(n))
                        .ToList();

                    if (drop.Count > 0)
                    {
                        try { await be.TryDropColumns(table, drop); }
                        catch (Exception ex) { XqlLog.Warn("DropColumns skipped: " + ex.Message); }
                        finally
                        {
                            // ✅ 헤더 기준으로 불필요 컬럼 삭제됨 → 캐시 무효화
                            XqlSheetView.InvalidateHeaderCache(ws.Name);
                        }
                    }
                }

                try { XqlSheetView.RegisterTableSheet(table, ws.Name); } catch { /* ignore */ }
            }
            catch (Exception ex) { XqlLog.Warn("EnsureActiveSheetSchema: " + ex.Message); }
            finally { XqlCommon.ReleaseCom(header, ws); }
        }
    }
}
