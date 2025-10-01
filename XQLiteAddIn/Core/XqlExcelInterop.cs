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
        private readonly object _uiGate = new();

        private static int _lastWarnTick;

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
                await EnsureActiveSheetSchemaAsync();
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

                var data = header.Offset[1, 0];                   // 헤더 아래 데이터 영역 시작
                var intersect = sh.Application.Intersect(target, data) as Excel.Range;
                if (intersect == null) return;                    // 헤더/외곽은 무시

                var table = string.IsNullOrWhiteSpace(sm.TableName) ? sh.Name : sm.TableName!;
                var keyColName = string.IsNullOrWhiteSpace(sm.KeyColumn) ? "id" : sm.KeyColumn!;

                foreach (Excel.Range cell in intersect.Cells)
                {
                    var colName = (string?)((header.Cells[1, cell.Column - header.Column + 1] as Excel.Range)!.Value2) ?? "";
                    if (string.IsNullOrWhiteSpace(colName)) colName = XqlCommon.ColumnIndexToLetter(cell.Column);

                    // key 셀 위치
                    var keyCell = sh.Cells[cell.Row, header.Column] as Excel.Range; // 헤더 첫 컬럼이 key 라는 가정
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
                }
            }
            catch (Exception ex)
            {
                XqlLog.Warn("OnWorksheetChange: " + ex.Message);
            }
        }
        /// <summary>
        /// 활성 시트의 헤더/메타를 읽어 서버에 테이블 없으면 생성, 누락 컬럼을 추가.
        /// </summary>
        private async Task EnsureActiveSheetSchemaAsync()
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
                                ColumnKind.Int => "integer",
                                ColumnKind.Real => "real",
                                ColumnKind.Bool => "boolean",
                                ColumnKind.Date => "integer",
                                _ => "text"
                            },
                            NotNull = !ct.Nullable,
                            Check = null
                        };
                    });
                    await be.TryAddColumns(table, defs);
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
                    }
                }
            }
            catch (Exception ex) { XqlLog.Warn("EnsureActiveSheetSchemaAsync: " + ex.Message); }
            finally { XqlCommon.ReleaseCom(header); XqlCommon.ReleaseCom(ws); }
        }
#if false
        private static void ApplyValidationVisual(Excel.Worksheet ws, Excel.Range cell, ValidationResult vr)
        {

            if (vr.IsOk)
            {
                SafeClearComment(cell);
                XqlSheetView.TryClearInvalidMark(cell); // ← 추가: 연분홍 제거
                return;
            }


            SafeClearComment(cell);
            SafeSetComment(cell, TruncateForComment(vr.Message));


            bool isDataCell = false;

            // 1) 표 바디에 속하면 무조건 데이터 셀
            var lo = cell.ListObject ?? XqlSheet.FindListObjectContaining(ws, cell);
            Excel.Range? inter = null;
            try
            {
                if (lo?.DataBodyRange != null)
                {
                    inter = XqlCommon.IntersectSafe(ws, lo.DataBodyRange, cell);
                    isDataCell = inter != null;
                }

                // 2) 표가 없으면: 헤더 마커(or Fallback) 아래 행부터를 데이터로 간주
                if (!isDataCell)
                {
                    Excel.Range? hdr = null;
                    try
                    {
                        if (!XqlSheet.TryGetHeaderMarker(ws, out hdr))
                            hdr = XqlSheet.GetHeaderRange(ws);
                        if (hdr != null) isDataCell = cell.Row > hdr.Row;
                    }
                    finally { XqlCommon.ReleaseCom(hdr); }
                }

                if (isDataCell)
                {
                    // 로그 + 에러 하이라이트
                    XqlLog.Warn(vr.Message, ws.Name, cell.Address[false, false]);
                    XqlSheetView.MarkInvalidCell(cell);

                    if (ShouldWarnNow())
                    {
                        System.Windows.Forms.MessageBox.Show(vr.Message, "XQLite",
                            System.Windows.Forms.MessageBoxButtons.OK,
                            System.Windows.Forms.MessageBoxIcon.Warning);
                    }
                }
            }
            catch { }
            finally { XqlCommon.ReleaseCom(inter); XqlCommon.ReleaseCom(lo); }
        }
#endif

        private static string TruncateForComment(string s)
        {
            if (string.IsNullOrEmpty(s)) return "";
            // 댓글 텍스트는 과도하게 길면 렌더링 문제가 있을 수 있음
            return s.Length <= 512 ? s : s.Substring(0, 509) + "...";
        }

        private static void SafeClearComment(Excel.Range cell)
        {
            try
            {
                var c = cell.Comment;
                if (c != null) c.Delete();
            }
            catch { /* 일부 워크시트 보호 등으로 실패 가능 */ }
        }

        private static void SafeSetComment(Excel.Range cell, string text)
        {
            try
            {
                cell.AddComment(text);
                if (cell.Comment != null)
                {
                    // 기본 비가시 상태
                    cell.Comment.Visible = false;
                    // 색상/스타일 등 꾸미기가 필요하면 여기서
                }
            }
            catch
            {
                // 실패 시 무시
            }
        }

        // ========= 컨텍스트/값 읽기 =========

        private static (string table, object rowKey, string colName, object? value)
        ReadCellContext(Excel.Worksheet sh, Excel.Range cell)
        {
            string tableName = sh.Name;
            Excel.Range? hc = null, kc = null;

            // 1) 표/헤더 찾기
            var lo = cell.ListObject ?? XqlSheet.FindListObjectContaining(sh, cell);
            Excel.Range? header = lo?.HeaderRowRange;
            if (header == null && XqlSheet.TryGetHeaderMarker(sh, out var hdr))
                header = hdr;

            if (header != null)
            {
                // 표가 있으면 테이블명 매핑 사용
                if (lo != null)
                    tableName = XqlTableNameMap.Map(lo.Name, sh.Name);

                // 컬럼명: 현재 셀의 '헤더 열'에서 읽기 (비었으면 A/B/C…)
                hc = (Excel.Range)sh.Cells[header.Row, cell.Column];
                var name = (hc.Value2 as string)?.Trim();
                if (string.IsNullOrEmpty(name)) name = XqlCommon.ColumnIndexToLetter(hc.Column);
                XqlCommon.ReleaseCom(hc);
                string colName = name!;

                // 키: 메타 키 우선, 없으면 헤더 기준 첫 컬럼
                var meta = XqlAddIn.Sheet!.GetOrCreateSheet(sh.Name);
                int keyAbsCol = XqlSheet.FindKeyColumnAbsolute(header, meta.KeyColumn);
                var rkCell = (Excel.Range)sh.Cells[cell.Row, keyAbsCol];
                object rowKey = rkCell.Value2 ?? cell.Row;
                XqlCommon.ReleaseCom(rkCell);

                return (tableName, rowKey, colName, ReadCellValueNormalized(cell));
            }

            // 2) 폴백: 기존 동작 유지
            string fallbackCol;
            object fallbackKey;

            try
            {
                hc = (Excel.Range)sh.Cells[1, cell.Column];
                var hName = (hc.Value2 as string)?.Trim();
                fallbackCol = string.IsNullOrEmpty(hName)
                     ? XqlCommon.ColumnIndexToLetter(cell.Column)
                     : hName!;
                kc = (Excel.Range)sh.Cells[cell.Row, 1];
                fallbackKey = kc.Value2 ?? cell.Row;
            }
            finally { XqlCommon.ReleaseCom(hc); XqlCommon.ReleaseCom(kc); }
            return (tableName, fallbackKey, fallbackCol, ReadCellValueNormalized(cell));
        }


        private static object? ReadCellValueNormalized(Excel.Range cell)
        {
            // Excel Value2는 Date를 OADate(double)로 반환.
            var v = cell.Value2;
            if (v == null) return null;

            // 배열/Range는 단일 셀에서 나오지 않아야 하나, 안전망
            if (v is Array) return v;

            // 숫자/날짜 구분: 날짜 포맷 여부를 이용하거나 OADate 변환 시도
            if (v is double d)
            {
                try
                {
                    // 강제 날짜 변환은 하지 않고, 소비자가 필요 시 판단.
                    // 여기서는 double 그대로 보낸다.
                    return d;
                }
                catch
                {
                    return d;
                }
            }

            if (v is string s) return s;
            if (v is bool b) return b;

            return v; // 그 외는 원본 유지
        }


        // 유효성 경고 쓰로틀 (연속 MsgBox 억제)
        private static bool ShouldWarnNow()
        {
            int now = Environment.TickCount;
            // 300ms 이내 재호출이면 무시
            if (unchecked(now - _lastWarnTick) < 300) return false;
            _lastWarnTick = now;
            return true;
        }

        // ========= UI 스레드 보조 =========

        private void RunOnUiThread(Action action)
        {
            // ExcelDna는 호출 스레드가 종종 UI 스레드가 아닐 수 있다.
            lock (_uiGate)
            {
                try
                {
                    ExcelAsyncUtil.QueueAsMacro(() =>
                    {
                        try { action(); } catch { /* swallow */ }
                    });
                }
                catch
                {
                    try { action(); } catch { /* swallow */ }
                }
            }
        }
    }
}
