// XqlCollab.cs (Migration + RelativeKey 내장 통합판)
using System;
using System.Text.RegularExpressions;
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
            } catch { }
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
                var key = LockKeyMigration.MigrateIfNeeded(app, resourceKey);
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
            try
            {
                var app = (Excel.Application)ExcelDnaUtil.Application;
                var rng = (Excel.Range)app.Selection;
                if (rng == null) return false;

                var ws = (Excel.Worksheet)rng.Worksheet;
                var lo = rng.ListObject ?? XqlSheet.FindListObjectContaining(ws, rng);
                if (lo?.HeaderRowRange == null) return false;

                var tableName = XqlTableNameMap.Map(lo.Name, ws.Name);
                var (header, headers) = XqlSheet.GetHeaderAndNames(ws);
                // 선택한 셀의 헤더 인덱스 계산
                int colIndex = rng.Column - lo.HeaderRowRange.Column; // 0-base
                if (colIndex < 0 || colIndex >= headers.Count) return false;

                string headerName = headers[colIndex];
                int hRow = lo.HeaderRowRange.Row;
                int hCol = lo.HeaderRowRange.Column;
                int colOffset = colIndex; // 헤더 좌상단 기준 상대 offset(0-base)

                var key = XqlSheet.ColumnKey(ws.Name, tableName, hRow, hCol, colOffset, headerName);
                await _backend.AcquireLock(key, _nickname).ConfigureAwait(false);
                return true;
            }
            catch { return false; }
        }

        /// <summary>현재 선택한 셀을 상대키로 계산해 획득</summary>
        public async Task<bool> AcquireCurrentCell()
        {
            try
            {
                var app = (Excel.Application)ExcelDnaUtil.Application;
                var rng = (Excel.Range)app.Selection;
                if (rng == null) return false;

                var ws = (Excel.Worksheet)rng.Worksheet;
                var lo = rng.ListObject ?? XqlSheet.FindListObjectContaining(ws, rng);
                if (lo?.HeaderRowRange == null) return false;

                var tableName = XqlTableNameMap.Map(lo.Name, ws.Name);
                int hRow = lo.HeaderRowRange.Row;
                int hCol = lo.HeaderRowRange.Column;

                int rowOffset = rng.Row - (hRow + 1); // 데이터 첫행 = header+1
                int colOffset = rng.Column - hCol;

                var key = XqlSheet.CellKey(ws.Name, tableName, hRow, hCol, rowOffset, colOffset);
                await _backend.AcquireLock(key, _nickname).ConfigureAwait(false);
                return true;
            }
            catch { return false; }
        }

        // (선택) 특정 구키를 새키로 서버에서 교체 시도: 새키 획득 후 내 락 해제
        public async Task<bool> MigrateOnServer(string oldKey)
        {
            try
            {
                var app = (Excel.Application)ExcelDnaUtil.Application;
                var newKey = LockKeyMigration.MigrateIfNeeded(app, oldKey);
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
            try
            {
                var app = (Excel.Application)ExcelDnaUtil.Application;
                var rng = (Excel.Range)app.Selection;
                if (rng == null) return null;

                var ws = (Excel.Worksheet)rng.Worksheet;
                var lo = rng.ListObject ?? XqlSheet.FindListObjectContaining(ws, rng);
                if (lo?.HeaderRowRange == null) return null;

                var tableName = XqlTableNameMap.Map(lo.Name, ws.Name);
                int hRow = lo.HeaderRowRange.Row;
                int hCol = lo.HeaderRowRange.Column;

                int rowOffset = rng.Row - (hRow + 1);
                int colOffset = rng.Column - hCol;

                return XqlSheet.CellKey(ws.Name, tableName, hRow, hCol, rowOffset, colOffset);
            }
            catch { return null; }
        }

        // ─────────────────────────────────────────────────────────────────────
        // 마이그레이션(구 포맷 → 상대키)  ─ 내부 클래스로 내장
        // ─────────────────────────────────────────────────────────────────────
        private static class LockKeyMigration
        {
            // 예: cell:Sheet!B5  | cell:Sheet!$C$10
            private static readonly Regex RxOldCell =
                new(@"^cell:(?<sheet>[^!]+)!(?<addr>\$?[A-Z]+\$?\d+)$",
                    RegexOptions.Compiled | RegexOptions.CultureInvariant | RegexOptions.IgnoreCase);

            // 예: column:Table.Column  | col:Table.Column
            private static readonly Regex RxOldColumn =
                new(@"^(?:column|col):(?<table>[^\.]+)\.(?<column>.+)$",
                    RegexOptions.Compiled | RegexOptions.CultureInvariant | RegexOptions.IgnoreCase);

            public static string MigrateIfNeeded(Excel.Application app, string oldKey)
            {
                if (string.IsNullOrWhiteSpace(oldKey)) return oldKey;

                // 1) cell:Sheet!A1
                var mCell = RxOldCell.Match(oldKey);
                if (mCell.Success)
                {
                    var sheet = mCell.Groups["sheet"].Value;
                    var addr = mCell.Groups["addr"].Value;

                    var ws = XqlSheet.FindWorksheet(app, sheet);
                    if (ws == null) return oldKey;

                    Excel.Range? rng = null; Excel.ListObject? lo = null;
                    try
                    {
                        rng = ws.Range[addr];
                        lo = rng?.ListObject ?? XqlSheet.FindListObjectContaining(ws, rng!);
                        if (lo?.HeaderRowRange == null) return oldKey;

                        int hRow = lo.HeaderRowRange.Row;
                        int hCol = lo.HeaderRowRange.Column;

                        int rowOffset = rng!.Row - (hRow + 1);
                        int colOffset = rng!.Column - hCol;

                        var tableName = XqlTableNameMap.Map(lo.Name, ws.Name);
                        return XqlSheet.CellKey(ws.Name, tableName, hRow, hCol, rowOffset, colOffset);
                    }
                    catch { return oldKey; }
                    finally { XqlCommon.ReleaseCom(lo); XqlCommon.ReleaseCom(rng); }
                }

                // 2) column:Table.Column
                var mCol = RxOldColumn.Match(oldKey);
                if (mCol.Success)
                {
                    var table = mCol.Groups["table"].Value;
                    var col = mCol.Groups["column"].Value;

                    try
                    {
                        var app2 = (Excel.Application)ExcelDnaUtil.Application;
                        var ws = (Excel.Worksheet)app2.ActiveSheet;
                        if (ws == null) return oldKey;

                        var lo = XqlSheet.FindListObjectByTable(ws, table);
                        if (lo?.HeaderRowRange == null) return oldKey;

                        var (header, headers) = XqlSheet.GetHeaderAndNames(ws);
                        int colIndex = headers.FindIndex(h => string.Equals(h, col, StringComparison.Ordinal));
                        if (colIndex < 0) return oldKey;

                        int hRow = lo.HeaderRowRange.Row;
                        int hCol = lo.HeaderRowRange.Column;
                        int headerCol = lo.HeaderRowRange.Column + colIndex;
                        int colOffset = headerCol - hCol;

                        return XqlSheet.ColumnKey(ws.Name, table, hRow, hCol, colOffset, col);
                    }
                    catch { return oldKey; }
                }

                // 3) 이미 신 포맷/미인식 포맷
                return oldKey;
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
