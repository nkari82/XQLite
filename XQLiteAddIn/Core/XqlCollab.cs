// XqlCollab.cs (UI-thread marshaling via XqlCommon.OnExcelThreadAsync, no blocking waits)
using ExcelDna.Integration;
using System;
using System.Threading;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using Timer = System.Threading.Timer;

namespace XQLite.AddIn
{
    internal sealed class XqlCollab : IDisposable
    {
        private readonly IXqlBackend _backend;
        private readonly string _nickname;
        private readonly Timer _refresh;

        private long _lastPresenceMs;
        private string? _lastPresenceSig;

        public XqlCollab(IXqlBackend backend, string nickname, int refreshSec = 3)
        {
            _backend = backend ?? throw new ArgumentNullException(nameof(backend));
            _nickname = string.IsNullOrWhiteSpace(nickname) ? "anonymous" : nickname.Trim();

            _refresh = new Timer(async _ => await SafeRefresh().ConfigureAwait(false), null, Timeout.Infinite, Timeout.Infinite);
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

        // 선택 변경 시 즉시 presence 디바운스 전송
        public void SelectionChanged(string sheet, string cellAddr)
        {
            try
            {
                var now = XqlCommon.NowMs();
                if (now - _lastPresenceMs < 800) return; // 0.8s 디바운스

                var sig = $"{sheet}|{cellAddr}";
                if (sig == _lastPresenceSig) return;

                _lastPresenceSig = sig;
                _lastPresenceMs = now;

                var be = XqlAddIn.Backend;
                var nick = XqlConfig.Nickname ?? Environment.UserName;
                if (be == null || string.IsNullOrWhiteSpace(nick)) return;

                // 비차단 fire-and-forget (백엔드 호출은 UI 스레드 불필요)
                _ = be.PresenceTouch(nick, sheet, cellAddr);
            }
            catch { /* non-fatal */ }
        }

        // ─────────────────────────────────────────────────────────────────────
        // Presence
        // ─────────────────────────────────────────────────────────────────────
        private async Task SafeRefresh()
        {
            try
            {
                // UI 스냅샷만 Excel 스레드에서 짧게 획득
                var (sheet, cell) = await TryGetCurrentSheetAndCellKeyOrNullAsync().ConfigureAwait(false);

                // 네트워크 호출은 백그라운드에서
                await _backend.PresenceTouch(_nickname, sheet, cell).ConfigureAwait(false);
            }
            catch
            {
                // 네트워크/일시 오류 무시
            }
        }

        // ─────────────────────────────────────────────────────────────────────
        // Lock APIs (상대키 기반)
        // ─────────────────────────────────────────────────────────────────────
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
            try
            {
                await _backend.ReleaseLocksBy(_nickname).ConfigureAwait(false);
                return true;
            }
            catch { return false; }
        }

        /// <summary>현재 선택의 컬럼을 상대키로 계산해 획득</summary>
        public async Task<bool> AcquireCurrentColumn()
        {
            // 1) UI 스냅샷(키 계산)만 Excel 스레드에서
            var key = await XqlCommon.OnExcelThreadAsync<string?>(() =>
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
                finally { XqlCommon.ReleaseCom(headerCell, rng, lo, ws); }
            }).ConfigureAwait(false);

            if (string.IsNullOrEmpty(key)) return false;

            // 2) 서버 호출은 백그라운드
            try
            {
                await _backend.AcquireLock(key!, _nickname).ConfigureAwait(false);
                return true;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>현재 선택한 셀을 상대키로 계산해 획득</summary>
        public async Task<bool> AcquireCurrentCell()
        {
            // 1) UI 스냅샷(키 계산)만 Excel 스레드에서
            var key = await XqlCommon.OnExcelThreadAsync<string?>(() =>
            {
                Excel.Range? rng = null; Excel.Worksheet? ws = null; Excel.ListObject? lo = null;
                try
                {
                    var app = (Excel.Application)ExcelDnaUtil.Application;
                    rng = app.Selection as Excel.Range; if (rng == null) return null;

                    ws = (Excel.Worksheet)rng.Worksheet;
                    lo = rng.ListObject ?? XqlSheet.FindListObjectContaining(ws, rng);
                    if (lo?.HeaderRowRange == null) return null;

                    int hRow = lo.HeaderRowRange.Row, hCol = lo.HeaderRowRange.Column;
                    int dr = rng.Row - (hRow + 1), dc = rng.Column - hCol;
                    return XqlSheet.CellKey(ws.Name, XqlTableNameMap.Map(lo.Name, ws.Name), hRow, hCol, dr, dc);
                }
                finally
                {
                    XqlCommon.ReleaseCom(rng, lo, ws);
                }
            }).ConfigureAwait(false);

            if (string.IsNullOrEmpty(key)) return false;

            // 2) 서버 호출은 백그라운드
            try
            {
                await _backend.AcquireLock(key!, _nickname).ConfigureAwait(false);
                return true;
            }
            catch
            {
                return false;
            }
        }

        // UI가 현재 커서를 Presence에 태그하고 싶을 때 사용 (UI 스냅샷 전용)
        private static Task<(string? sheet, string? cellKey)> TryGetCurrentSheetAndCellKeyOrNullAsync()
        {
            return XqlCommon.OnExcelThreadAsync<(string?, string?)>(() =>
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
                finally { XqlCommon.ReleaseCom(rng, lo, ws); }
            });
        }

        // (선택) 키를 더블클릭으로 점프할 때 사용: 상대키 → 현재 Range 복원
        public static bool TryJumpTo(string key)
        {
            // 이 API는 시그니처를 유지(동기). 내부적으로 UI 스레드에서만 실행.
            try
            {
                var ok = XqlCommon.OnExcelThreadAsync<bool>(() =>
                {
                    var app = (Excel.Application)ExcelDnaUtil.Application;
                    if (XqlSheet.TryParse(key, out var desc) &&
                        XqlSheet.TryResolve(app, desc, out var range, out _, out _))
                    {
                        try { range?.Select(); return true; }
                        finally { XqlCommon.ReleaseCom(range); }
                    }
                    return false;
                }).GetAwaiter().GetResult(); // 호출측이 비-UI일 가능성이 높아 동기 대기 허용
                return ok;
            }
            catch
            {
                return false;
            }
        }
    }
}
