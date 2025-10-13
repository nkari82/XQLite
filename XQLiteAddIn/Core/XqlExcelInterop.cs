// XqlExcelInterop.cs
using ExcelDna.Integration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
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
    internal sealed class XqlExcelInterop : IDisposable
    {
        private readonly Excel.Application _app;
        private readonly XqlSync _sync;
        private readonly XqlCollab _collab;
        private readonly XqlSheet _sheet;
        private readonly XqlBackup _backup;

        private bool _started;

        public XqlExcelInterop(Excel.Application app, XqlSync sync, XqlCollab collab, XqlSheet sheet, XqlBackup backup)
        {
            _app = app ?? throw new ArgumentNullException(nameof(app));
            _sync = sync ?? throw new ArgumentNullException(nameof(sync));
            _collab = collab ?? throw new ArgumentNullException(nameof(collab));
            _sheet = sheet ?? throw new ArgumentNullException(nameof(sheet));
            _backup = backup ?? throw new ArgumentNullException(nameof(backup));
        }

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
            _app.WorkbookOpen += App_WorkbookOpen;
            _app.WorkbookBeforeClose += App_WorkbookBeforeClose;
        }

        public void Stop()
        {
            if (!_started) return;
            _started = false;

            _app.SheetChange -= App_SheetChange;
            _app.WorkbookOpen -= App_WorkbookOpen;
            _app.WorkbookBeforeClose -= App_WorkbookBeforeClose;
        }

        public void Dispose()
        {
            Stop();
        }

        // ========= 명령(리본/메뉴) =========
        public async void Cmd_CommitSmart()
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
                throw;
            }
        }


        public async void Cmd_PullOnly()
        {
            try { await _sync.PullSince(); }
            catch (Exception ex)
            {
                XqlLog.Warn("PullOnly failed: " + ex.Message);
                throw;
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
            // 필요 시 통합문서 메타 초기화 등
            XqlCommon.ReleaseCom(wb);
        }

        private void App_WorkbookBeforeClose(Excel.Workbook wb, ref bool Cancel)
        {
            // 락 해제, 프레즌스 정리 등
            var _ = _collab.ReleaseByMe();
            XqlCommon.ReleaseCom(wb);
        }

        // 변경 이벤트에서 호출
        private void App_SheetChange(object Sh, Excel.Range target)
        {
            Excel.Worksheet? sh = null;
            try
            {
                sh = Sh as Excel.Worksheet;
                if (sh == null)
                    return;

                var sm = _sheet.GetOrCreateSheet(sh.Name);
                var header = XqlSheetView.ResolveHeader(sh, null, _sheet);      // 헤더 Range
                if (header == null) return;

                // ✅ 헤더가 수정된 경우: 캐시 무효화 & 데이터 큐잉은 하지 않음
                try
                {
                    var hitHeader = sh.Application.Intersect(target, header) as Excel.Range;
                    if (hitHeader != null)
                    {
                        XqlCommon.ReleaseCom(hitHeader);
                        XqlSheetView.InvalidateHeaderCache(sh.Name);
                        return;
                    }
                }
                catch { /* ignore */ }


                Excel.Range? data = null;
                Excel.Range? intersect = null;
                Excel.ListObject? lo = null;
                try
                {
                    // 표가 있으면 DataBodyRange, 없으면 헤더 폭만큼 시트 끝까지
                    lo = XqlSheet.FindListObjectContaining(sh, header);
                    if (lo?.DataBodyRange != null)
                        data = lo.DataBodyRange;
                    else
                    {
                        var first = (Excel.Range)header.Offset[1, 0];
                        var last = sh.Cells[sh.Rows.Count, header.Column + header.Columns.Count - 1];
                        data = sh.Range[first, last];
                        XqlCommon.ReleaseCom(first);
                        XqlCommon.ReleaseCom(last);
                    }
                    intersect = sh.Application.Intersect(target, data) as Excel.Range;
                }
                finally
                {
                    // intersect는 아래에서 사용하므로 여기서 해제하지 않음
                }

                if (intersect == null) return;                    // 헤더/외곽은 무시

                var table = string.IsNullOrWhiteSpace(sm.TableName) ? sh.Name : sm.TableName!;
                var keyColName = string.IsNullOrWhiteSpace(sm.KeyColumn) ? "id" : sm.KeyColumn!;

                foreach (Excel.Range cell in intersect.Cells)
                {
                    var colName = (string?)((header.Cells[1, cell.Column - header.Column + 1] as Excel.Range)!.Value2) ?? "";
                    if (string.IsNullOrWhiteSpace(colName)) colName = XqlCommon.ColumnIndexToLetter(cell.Column);

                    // key 셀 위치
                    int keyAbsCol = XqlSheet.FindKeyColumnAbsolute(header, sm.KeyColumn);
                    var keyCell = sh.Cells[cell.Row, keyAbsCol] as Excel.Range;
                    var rowKeyObj = keyCell?.Value2;
                    string? rowKey = rowKeyObj?.ToString();

                    // 새 행이면 id 생성 (텍스트 키 가정) — 프로젝트 정책에 맞게 교체 가능
                    if (string.IsNullOrWhiteSpace(rowKey))
                    {
                        if (!string.Equals(colName, keyColName, StringComparison.OrdinalIgnoreCase))
                            continue; // 키가 없고 key가 아닌 셀 수정이면 보류

                        rowKey = Guid.NewGuid().ToString("N").Substring(0, 12);
                        keyCell!.Value2 = rowKey; // 시트에 즉시 반영
                    }

                    // 값 뽑기 (JSON/날짜/불린 변환은 기존 유틸 사용 가능)
                    object? value = cell.Value2;

                    // ⬇️ 딱 이 셀만 큐잉
                    _sync.EnqueueIfChanged(table, rowKey!, colName, value);
                    XqlCommon.ReleaseCom(keyCell);
                }

                XqlCommon.ReleaseCom(intersect);
                XqlCommon.ReleaseCom(data);
                XqlCommon.ReleaseCom(lo);
            }
            catch (Exception ex)
            {
                XqlLog.Warn("OnWorksheetChange: " + ex.Message);
            }
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
                                XqlSheet.ColumnKind.Bool => "boolean",
                                XqlSheet.ColumnKind.Date => "integer",
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
            }
            catch (Exception ex) { XqlLog.Warn("EnsureActiveSheetSchema: " + ex.Message); }
            finally { XqlCommon.ReleaseCom(header); XqlCommon.ReleaseCom(ws); }
        }
    }
}
