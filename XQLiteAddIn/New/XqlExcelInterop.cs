// XqlExcelInterop.cs
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using ExcelDna.Integration;
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
        private readonly XqlMetaRegistry _meta;

        private bool _started;
        private readonly object _uiGate = new();

        // 선택 변경 스팸 억제를 위한 디바운서
        private readonly System.Threading.Timer _heartbeatDebounce;
        private volatile string _lastCellRef = string.Empty;

        public XqlExcelInterop(Excel.Application app, XqlSync sync, XqlCollab collab, XqlMetaRegistry meta)
        {
            _app = app ?? throw new ArgumentNullException(nameof(app));
            _sync = sync ?? throw new ArgumentNullException(nameof(sync));
            _collab = collab ?? throw new ArgumentNullException(nameof(collab));
            _meta = meta ?? throw new ArgumentNullException(nameof(meta));

            _heartbeatDebounce = new System.Threading.Timer(_ => SendHeartbeat(), null, Timeout.Infinite, Timeout.Infinite);
        }

        // ========= 수명 주기 =========

        public void Start()
        {
            if (_started) return;
            _started = true;

            _app.SheetChange += App_SheetChange;
            _app.SheetSelectionChange += App_SheetSelectionChange;
            _app.WorkbookOpen += App_WorkbookOpen;
            _app.WorkbookBeforeClose += App_WorkbookBeforeClose;

            // 최초 하트비트 타이머 기동(사용자 선택이 들어오면 BYHOUR=2초 디바운스)
            _heartbeatDebounce.Change(Timeout.Infinite, Timeout.Infinite);
        }

        public void Stop()
        {
            if (!_started) return;
            _started = false;

            _app.SheetChange -= App_SheetChange;
            _app.SheetSelectionChange -= App_SheetSelectionChange;
            _app.WorkbookOpen -= App_WorkbookOpen;
            _app.WorkbookBeforeClose -= App_WorkbookBeforeClose;

            _heartbeatDebounce.Change(Timeout.Infinite, Timeout.Infinite);
        }

        public void Dispose()
        {
            Stop();
            _heartbeatDebounce.Dispose();
        }

        // ========= 명령(리본/메뉴) =========

        public void Cmd_CommitSync()
        {
            // 서버에서 증분 Pull → Excel 반영은 XqlSync가 수행 (머지/충돌 로직 포함)
            _ = _sync.PullSince(_sync.MaxRowVersion);
        }

        public void Cmd_RecoverFromExcel()
        {
            // 원클릭 복구: 엑셀 파일을 원본으로 DB 재생성
            // #FIXME Backup
#if false
            XqlAddIn.Backup?.RecoverFromExcel();
#endif
        }

        public void Cmd_SetHeaderTooltipsForActiveSheet()
        {
            RunOnUiThread(() =>
            {
                var sh = (Excel.Worksheet)_app.ActiveSheet;
                if (sh == null) return;

                if (!_meta.TryGetSheet(sh.Name, out var sheetMeta))
                    return;

                var dict = _meta.TryGetSheet(sh.Name, out var sm)
                    ? sm.Columns.ToDictionary(
                        kv => kv.Key,
                        kv => kv.Value.ToTooltip(),
                        StringComparer.Ordinal)
                    : new Dictionary<string, string>(StringComparer.Ordinal);
                SetHeaderTooltips(sh, dict);
                ReleaseCom(sh);
            });
        }

        // ========= Excel 이벤트 =========

        private void App_WorkbookOpen(Excel.Workbook wb)
        {
            // 필요 시 통합문서 메타 초기화 등
            ReleaseCom(wb);
        }

        private void App_WorkbookBeforeClose(Excel.Workbook wb, ref bool Cancel)
        {
            // 락 해제, 프레즌스 정리 등
            var nickname = XqlAddIn.Cfg?.Nickname ?? "anonymous";
            _collab.ReleaseLocksBy(nickname);
            ReleaseCom(wb);
        }

        private void App_SheetSelectionChange(object Sh, Excel.Range Target)
        {
            try
            {
                var sh = Sh as Excel.Worksheet;
                if (sh == null || Target == null) return;

                var cellRef = $"{sh.Name}!{Target.Address[false, false]}";
                _lastCellRef = cellRef;

                // 2초 디바운스 하트비트
                _heartbeatDebounce.Change(2000, Timeout.Infinite);

                ReleaseCom(Target);
                ReleaseCom(sh);
            }
            catch { /* swallow */ }
        }

        private void App_SheetChange(object Sh, Excel.Range Target)
        {
            try
            {
                var sh = Sh as Excel.Worksheet;
                if (sh == null || Target == null) return;

                // 변경 범위가 여러 셀일 수 있음
                foreach (Excel.Range cell in Target.Cells)
                {
                    try
                    {
                        var ctx = ReadCellContext(sh, cell); // table,rowKey,colName,value
                        var vr = _meta.ValidateCell(sh.Name, ctx.colName, ctx.value);
                        ApplyValidationVisual(cell, vr);

                        if (vr.IsOk)
                        {
                            _sync.EnqueueCellEdit(ctx.table, ctx.rowKey, ctx.colName, ctx.value);
                        }
                    }
                    catch (Exception ex)
                    {
                        // 셀 주석으로 에러 표시
                        SafeSetComment(cell, $"Edit error: {ex.Message}");
                    }
                    finally
                    {
                        ReleaseCom(cell);
                    }
                }

                ReleaseCom(Target);
                ReleaseCom(sh);
            }
            catch { /* swallow */ }
        }

        // ========= 하트비트 & 락 =========

        private void SendHeartbeat()
        {
            try
            {
                var nickname = XqlAddIn.Cfg?.Nickname ?? "anonymous";
                var cellRef = _lastCellRef;
                _collab.Heartbeat(nickname, cellRef);

                // 선택 셀에 대한 락(낙관적으로 시도)
                if (!string.IsNullOrEmpty(cellRef))
                    _collab.TryAcquireCellLock(cellRef, nickname);
            }
            catch { /* swallow */ }
        }

        // ========= 헤더 범위 / 툴팁 / 주석 =========

        /// <summary>
        /// 헤더: 1행의 연속된 타이틀 영역으로 가정(빈 셀을 만나면 종료).
        /// </summary>
        public Excel.Range GetHeaderRange(Excel.Worksheet sh)
        {
            Excel.Range? lastCell = null;
            try
            {
                int col = 1;
                for (; col <= sh.UsedRange.Columns.Count; col++)
                {
                    var cell = (Excel.Range)sh.Cells[1, col];
                    var txt = (cell?.Value2 as string)?.Trim();
                    if (string.IsNullOrEmpty(txt))
                    {
                        ReleaseCom(cell);
                        break;
                    }
                    ReleaseCom(lastCell);
                    lastCell = cell;
                }

                var lastCol = Math.Max(1, (lastCell?.Column as int?) ?? 1);
                var rg = sh.Range[sh.Cells[1, 1], sh.Cells[1, lastCol]];
                ReleaseCom(lastCell);
                return rg;
            }
            catch
            {
                ReleaseCom(lastCell);
                // fallback: A1
                return (Excel.Range)sh.Cells[1, 1];
            }
        }

        /// <summary>
        /// 헤더 셀에 컬럼 타입/체크 제약을 툴팁(주석)으로 표시한다.
        /// </summary>
        public void SetHeaderTooltips(Excel.Worksheet sh, IReadOnlyDictionary<string, string> colToTip)
        {
            Excel.Range? header = null;
            try
            {
                header = GetHeaderRange(sh);
                foreach (Excel.Range cell in header.Cells)
                {
                    try
                    {
                        var colName = (cell.Value2 as string)?.Trim();
                        if (string.IsNullOrEmpty(colName)) continue;
                        if (!colToTip.TryGetValue(colName, out var tip)) continue;

                        SafeClearComment(cell);
                        SafeSetComment(cell, tip);
                    }
                    finally
                    {
                        ReleaseCom(cell);
                    }
                }
            }
            finally
            {
                ReleaseCom(header);
            }
        }

        private static void ApplyValidationVisual(Excel.Range cell, ValidationResult vr)
        {
            if (vr.IsOk)
            {
                SafeClearComment(cell);
                return;
            }

            SafeClearComment(cell);
            SafeSetComment(cell, TruncateForComment(vr.Message));
        }

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

        private static (string table, object rowKey, string colName, object? value) ReadCellContext(Excel.Worksheet sh, Excel.Range cell)
        {
            // 가정:
            // - 1행: 헤더(컬럼명)
            // - 1열: 키 컬럼(행 키)
            string table = sh.Name;

            string colName = ((Excel.Range)sh.Cells[1, cell.Column]).Value2 as string ?? $"C{cell.Column}";
            object rowKey = ((Excel.Range)sh.Cells[cell.Row, 1]).Value2 ?? cell.Row;

            object? value = ReadCellValueNormalized(cell);
            return (table, rowKey, colName, value);
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

        // ========= COM 해제 유틸 =========

        private static void ReleaseCom(object? com)
        {
            try
            {
                if (com != null && Marshal.IsComObject(com))
                    Marshal.FinalReleaseComObject(com);
            }
            catch { /* ignore */ }
        }
    }
}
