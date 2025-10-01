// XqlCollab.cs (Migration + RelativeKey 내장 통합판, refactored)
using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;
using System;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Input;
using Excel = Microsoft.Office.Interop.Excel;
using Timer = System.Threading.Timer;

namespace XQLite.AddIn
{
    internal sealed class XqlCollab : IDisposable
    {
        private readonly IXqlBackend _backend;
        private readonly string _nickname;
        private readonly Timer _refresh;

        public XqlCollab(IXqlBackend backend, string nickname, int refreshSec = 3)
        {
            _backend = backend ?? throw new ArgumentNullException(nameof(backend));
            _nickname = string.IsNullOrWhiteSpace(nickname) ? "anonymous" : nickname.Trim();
            _refresh = new Timer(async _ => await SafeRefresh(), null, Timeout.Infinite, Timeout.Infinite);
            _ = SafeRefresh(); // 즉시 1회
            _refresh.Change(TimeSpan.FromSeconds(refreshSec), TimeSpan.FromSeconds(refreshSec));
        }

        public void Dispose()
        {
            try
            {
                _refresh.Change(Timeout.Infinite, Timeout.Infinite);
                _refresh.Dispose();
            }
            catch { }
        }

        // ─────────────────────────────────────────────────────────────────────
        // Presence
        // ─────────────────────────────────────────────────────────────────────
        private async Task SafeRefresh()
        {
            if (_refresh == null)
                return;

            try
            {
                // 현재 선택 위치를 상대키로 얻고, sheet 이름도 함께 보냄
                var (sheet, cell) = TryGetCurrentSheetAndCellKeyOrNull();
                await _backend.PresenceTouch(_nickname, sheet, cell).ConfigureAwait(false);
            }
            catch { /* 네트워크 일시 오류는 무시 */ }
        }

        // ─────────────────────────────────────────────────────────────────────
        // Lock APIs (항상 상대 키로 정규화하여 서버 호출)
        // ─────────────────────────────────────────────────────────────────────
        // 🔧 교체: 입력 키를 그대로 사용 (상대키만 지원)
        public async Task<bool> Acquire(string resourceKey)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(resourceKey)) return false;
                await _backend.AcquireLock(resourceKey.Trim(), _nickname).ConfigureAwait(false);
                return true;
            }
            catch { return false; }
        }

        public async Task<bool> ReleaseByMe()
        {
            try { await _backend.ReleaseLocksBy(_nickname).ConfigureAwait(false); return true; }
            catch { return false; }
        }

        /// <summary>현재 선택의 컬럼을 상대키로 계산해 획득</summary>
        public async Task<bool> AcquireCurrentColumn()
        {
            // COM은 UI 스레드에서 계산 → 키만 받아와서 서버 호출
            var key = OnMainThread<string?>(() =>
            {
                Excel.Range? rng = null; Excel.Worksheet? ws = null; Excel.ListObject? lo = null; Excel.Range? headerCell = null;
                try
                {
                    var app = (Excel.Application)ExcelDnaUtil.Application;
                    rng = app.Selection as Excel.Range;
                    if (rng == null) return null;
                    ws = (Excel.Worksheet)rng.Worksheet;
                    lo = rng.ListObject ?? XqlSheet.FindListObjectContaining(ws, rng);
                    if (lo?.HeaderRowRange == null) return null;
                    // 선택 좌상단 기준 컬럼 인덱스
                    int colIndex = Math.Max(0, Math.Min(rng.Column - lo.HeaderRowRange.Column, lo.ListColumns.Count - 1));
                    headerCell = (Excel.Range)lo.HeaderRowRange.Cells[1, colIndex + 1];
                    var headerName = (headerCell.Value2 as string)?.Trim();
                    if (string.IsNullOrEmpty(headerName))
                        headerName = XqlCommon.ColumnIndexToLetter(headerCell.Column);
                    string tableName = XqlTableNameMap.Map(lo.Name, ws.Name);
                    int hRow = lo.HeaderRowRange.Row;
                    int hCol = lo.HeaderRowRange.Column;
                    int colOffset = colIndex;
                    return XqlSheet.ColumnKey(ws.Name, tableName, hRow, hCol, colOffset, headerName!);
                }
                finally { XqlCommon.ReleaseCom(headerCell); XqlCommon.ReleaseCom(lo); XqlCommon.ReleaseCom(ws); XqlCommon.ReleaseCom(rng); }
            });
            if (string.IsNullOrEmpty(key)) return false;
            try { await _backend.AcquireLock(key!, _nickname).ConfigureAwait(false); return true; }
            catch { return false; }
        }

        /// <summary>현재 선택한 셀을 상대키로 계산해 획득</summary>
        public async Task<bool> AcquireCurrentCell()
        {
            Excel.Range? rng = null; Excel.Worksheet? ws = null; Excel.ListObject? lo = null;
            try
            {
                var app = (Excel.Application)ExcelDnaUtil.Application;
                rng = app.Selection as Excel.Range; if (rng == null) return false;
                ws = (Excel.Worksheet)rng.Worksheet;
                lo = rng.ListObject ?? XqlSheet.FindListObjectContaining(ws, rng);
                if (lo?.HeaderRowRange == null) return false;

                int hRow = lo.HeaderRowRange.Row, hCol = lo.HeaderRowRange.Column;
                int dr = rng.Row - (hRow + 1), dc = rng.Column - hCol;
                var key = XqlSheet.CellKey(ws.Name, XqlTableNameMap.Map(lo.Name, ws.Name), hRow, hCol, dr, dc);
                await _backend.AcquireLock(key, _nickname).ConfigureAwait(false);
                return true;
            }
            catch
            {
                return false;
            }
            finally
            {
                XqlCommon.ReleaseCom(lo);
                XqlCommon.ReleaseCom(ws);
                XqlCommon.ReleaseCom(rng);
            }
        }

        // UI가 현재 커서를 Presence에 태그하고 싶을 때 사용
        private static (string? sheet, string? cellKey) TryGetCurrentSheetAndCellKeyOrNull()
        {
            return OnMainThread<(string?, string?)>(() =>
            {
                Excel.Range? rng = null; Excel.Worksheet? ws = null; Excel.ListObject? lo = null;
                try
                {
                    var app = (Excel.Application)ExcelDnaUtil.Application;
                    rng = app.Selection as Excel.Range;
                    if (rng == null) return (null, null);
                    ws = (Excel.Worksheet)rng.Worksheet;
                    lo = rng.ListObject ?? XqlSheet.FindListObjectContaining(ws, rng);
                    if (lo?.HeaderRowRange == null) return (ws?.Name, null);
                    string tableName = XqlTableNameMap.Map(lo.Name, ws.Name);
                    int hRow = lo.HeaderRowRange.Row;
                    int hCol = lo.HeaderRowRange.Column;
                    int rowOffset = rng.Row - (hRow + 1);
                    int colOffset = rng.Column - hCol;
                    return (ws.Name, XqlSheet.CellKey(ws.Name, tableName, hRow, hCol, rowOffset, colOffset));
                }
                finally { XqlCommon.ReleaseCom(lo); XqlCommon.ReleaseCom(ws); XqlCommon.ReleaseCom(rng); }
            })!;
        }

        // (선택) 키를 더블클릭으로 점프할 때 사용: 상대키 → 현재 Range 복원
        public static bool TryJumpTo(string key)
        {
            var ok = OnMainThread<bool?>(() =>
            {
                try
                {
                    var app = (Excel.Application)ExcelDnaUtil.Application;
                    if (XqlSheet.TryParse(key, out var desc) &&
                        XqlSheet.TryResolve(app, desc, out var range, out _, out _))
                    { range?.Select(); return true; }
                }
                catch { }
                return false;
            });
            return ok == true;
        }

        // UI(Excel) 스레드에서 안전하게 실행하고 결과를 돌려받는 헬퍼
        private static T? OnMainThread<T>(Func<T?> work, int timeoutMs = 800)
        {
            T? result = default;
            var done = new System.Threading.ManualResetEventSlim(false);
            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                try { result = work(); }
                catch { /* swallow – presence 실패는 무시 */ }
                finally { done.Set(); }
            });
            if (!done.Wait(timeoutMs)) return default; // Excel이 바쁠 때는 타임아웃 후 null
            return result;
        }
    }
}
