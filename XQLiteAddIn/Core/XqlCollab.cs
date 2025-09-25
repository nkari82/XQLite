// XqlCollab.cs (Migration + RelativeKey 내장 통합판, refactored)
using System;
using System.Threading;
using System.Threading.Tasks;
using ExcelDna.Integration;
using Excel = Microsoft.Office.Interop.Excel;

namespace XQLite.AddIn
{
    internal sealed class XqlCollab : IDisposable
    {
        private readonly IXqlBackend _backend;
        private readonly string _nickname;
        private readonly Timer _heartbeat;
        private volatile bool _started;

        public XqlCollab(IXqlBackend backend, string nickname, int heartbeatSec = 3)
        {
            _backend = backend ?? throw new ArgumentNullException(nameof(backend));
            _nickname = string.IsNullOrWhiteSpace(nickname) ? "anonymous" : nickname.Trim();
            _heartbeat = new Timer(async _ => await SafeHeartbeat(), null, Timeout.Infinite, Timeout.Infinite);
            _ = SafeHeartbeat(); // 즉시 1회
            _heartbeat.Change(TimeSpan.FromSeconds(heartbeatSec), TimeSpan.FromSeconds(heartbeatSec));
            _started = true;
        }

        public void Dispose()
        {
            _started = false;
            try
            {
                _heartbeat.Change(Timeout.Infinite, Timeout.Infinite);
                _heartbeat.Dispose();
            }
            catch { }
        }

        // ─────────────────────────────────────────────────────────────────────
        // Presence
        // ─────────────────────────────────────────────────────────────────────
        private async Task SafeHeartbeat()
        {
            if (!_started) return;
            try
            {
                var cell = TryGetCurrentRelativeCellKeyOrNull();
                await _backend.PresenceHeartbeat(_nickname, cell).ConfigureAwait(false);
            }
            catch { /* 네트워크 일시 오류는 무시 */ }
        }

        // ─────────────────────────────────────────────────────────────────────
        // Lock APIs (항상 상대 키로 정규화하여 서버 호출)
        // ─────────────────────────────────────────────────────────────────────
        public async Task<bool> Acquire(string resourceKey)
        {
            try
            {
                var app = (Excel.Application)ExcelDnaUtil.Application;
                var key = XqlSheet.MigrateLockKeyIfNeeded(app, resourceKey);
                await _backend.AcquireLock(key, _nickname).ConfigureAwait(false);
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
            Excel.Range? rng = null; Excel.Worksheet? ws = null; Excel.ListObject? lo = null; Excel.Range? headerCell = null;
            try
            {
                var app = (Excel.Application)ExcelDnaUtil.Application;
                rng = app.Selection as Excel.Range;
                if (rng == null) return false;

                ws = (Excel.Worksheet)rng.Worksheet;
                lo = rng.ListObject ?? XqlSheet.FindListObjectContaining(ws, rng);
                if (lo?.HeaderRowRange == null) return false;

                // 선택 좌상단 컬럼 → 표 헤더 기준 0-base index
                int colIndex = rng.Column - lo.HeaderRowRange.Column;
                if (colIndex < 0) colIndex = 0;
                if (lo.ListColumns != null && colIndex >= lo.ListColumns.Count)
                    colIndex = lo.ListColumns.Count - 1;
                if (colIndex < 0) return false;

                // 헤더 명은 반드시 표 헤더에서 읽는다(UsedRange 1행 X)
                headerCell = (Excel.Range)lo.HeaderRowRange.Cells[1, colIndex + 1];
                string? headerName = (headerCell.Value2 as string)?.Trim();
                if (string.IsNullOrEmpty(headerName))
                    headerName = XqlCommon.ColumnIndexToLetter(headerCell.Column);

                string tableName = XqlTableNameMap.Map(lo.Name, ws.Name);
                int hRow = lo.HeaderRowRange.Row;
                int hCol = lo.HeaderRowRange.Column;
                int colOffset = colIndex; // 헤더 좌상단 기준 0-base

                var key = XqlSheet.ColumnKey(ws.Name, tableName, hRow, hCol, colOffset, headerName);
                await _backend.AcquireLock(key, _nickname).ConfigureAwait(false);
                return true;
            }
            catch { return false; }
            finally
            {
                XqlCommon.ReleaseCom(headerCell);
                XqlCommon.ReleaseCom(lo);
                XqlCommon.ReleaseCom(ws);
                XqlCommon.ReleaseCom(rng);
            }
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
            catch { return false; }
            finally { XqlCommon.ReleaseCom(lo); XqlCommon.ReleaseCom(ws); XqlCommon.ReleaseCom(rng); }
        }

        // (선택) 특정 구키를 새키로 서버에서 교체 시도: 새키 획득 후 내 락 해제
        public async Task<bool> MigrateOnServer(string oldKey)
        {
            try
            {
                var app = (Excel.Application)ExcelDnaUtil.Application;
                var newKey = XqlSheet.MigrateLockKeyIfNeeded(app, oldKey);
                if (newKey == oldKey) return true; // 이미 신포맷

                await _backend.AcquireLock(newKey, _nickname).ConfigureAwait(false);
                await _backend.ReleaseLocksBy(_nickname).ConfigureAwait(false);
                return true;
            }
            catch { return false; }
        }

        // UI가 현재 커서를 Presence에 태그하고 싶을 때 사용
        private static string? TryGetCurrentRelativeCellKeyOrNull()
        {
            Excel.Range? rng = null; Excel.Worksheet? ws = null; Excel.ListObject? lo = null;
            try
            {
                var app = (Excel.Application)ExcelDnaUtil.Application;
                rng = app.Selection as Excel.Range;
                if (rng == null) return null;

                ws = (Excel.Worksheet)rng.Worksheet;
                lo = rng.ListObject ?? XqlSheet.FindListObjectContaining(ws, rng);
                if (lo?.HeaderRowRange == null) return null;

                string tableName = XqlTableNameMap.Map(lo.Name, ws.Name);
                int hRow = lo.HeaderRowRange.Row;
                int hCol = lo.HeaderRowRange.Column;

                int rowOffset = rng.Row - (hRow + 1);
                int colOffset = rng.Column - hCol;

                return XqlSheet.CellKey(ws.Name, tableName, hRow, hCol, rowOffset, colOffset);
            }
            catch { return null; }
            finally
            {
                XqlCommon.ReleaseCom(lo);
                XqlCommon.ReleaseCom(ws);
                XqlCommon.ReleaseCom(rng);
            }
        }

        // (선택) 키를 더블클릭으로 점프할 때 사용: 상대키 → 현재 Range 복원
        public static bool TryJumpTo(string key)
        {
            try
            {
                var app = (Excel.Application)ExcelDnaUtil.Application;
                if (XqlSheet.TryParse(key, out var desc) &&
                    XqlSheet.TryResolve(app, desc, out var range, out _, out _))
                {
                    range?.Select();
                    return true;
                }
            }
            catch { }
            return false;
        }
    }
}
