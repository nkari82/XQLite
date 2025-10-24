// XqlCollab.cs — SmartCom<T> 적용 (UI-thread marshaling via XqlCommon.OnExcelThreadAsync)
using ExcelDna.Integration;
using System;
using System.Threading;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using Timer = System.Threading.Timer;
using static XQLite.AddIn.XqlCommon;

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
                var now = NowMs();
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
            var key = await OnExcelThreadAsync<string?>(() =>
            {
                using var app = SmartCom<Excel.Application>.Wrap((Excel.Application)ExcelDnaUtil.Application);

                using var rng = SmartCom<Excel.Range>.Wrap(app.Value?.Selection as Excel.Range);
                if (rng?.Value == null) return null;

                using var ws = SmartCom<Excel.Worksheet>.Wrap((Excel.Worksheet)rng.Value.Worksheet);
                using var lo = SmartCom<Excel.ListObject>.Wrap(rng.Value.ListObject ?? XqlSheet.FindListObjectContaining(ws.Value!, rng.Value));
                if (lo?.Value?.HeaderRowRange == null) return null;

                // 선택 좌상단 기준 컬럼 인덱스
                int colIndex = Math.Max(0, Math.Min(rng.Value.Column - lo.Value.HeaderRowRange.Column, lo.Value.ListColumns.Count - 1));

                using var headerCell = SmartCom<Excel.Range>.Acquire(() => (Excel.Range)lo.Value.HeaderRowRange.Cells[1, colIndex + 1]);
                var headerName = (headerCell?.Value?.Value2 as string)?.Trim();
                if (string.IsNullOrEmpty(headerName))
                    headerName = ColumnIndexToLetter(headerCell!.Value!.Column);

                string tableName = XqlTableNameMap.Map(lo.Value.Name, ws.Value!.Name);
                int hRow = lo.Value.HeaderRowRange.Row;
                int hCol = lo.Value.HeaderRowRange.Column;
                int colOffset = colIndex;

                return XqlSheet.ColumnKey(ws.Value!.Name, tableName, hRow, hCol, colOffset, headerName!);
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
            var key = await OnExcelThreadAsync<string?>(() =>
            {
                using var app = SmartCom<Excel.Application>.Wrap((Excel.Application)ExcelDnaUtil.Application);
                using var rng = SmartCom<Excel.Range>.Wrap(app.Value?.Selection as Excel.Range);
                if (rng?.Value == null) return null;

                using var ws = SmartCom<Excel.Worksheet>.Wrap((Excel.Worksheet)rng.Value.Worksheet);
                using var lo = SmartCom<Excel.ListObject>.Wrap(rng.Value.ListObject ?? XqlSheet.FindListObjectContaining(ws.Value!, rng.Value));
                if (lo?.Value?.HeaderRowRange == null) return null;

                int hRow = lo.Value.HeaderRowRange.Row, hCol = lo.Value.HeaderRowRange.Column;
                int dr = rng.Value.Row - (hRow + 1), dc = rng.Value.Column - hCol;
                return XqlSheet.CellKey(ws.Value!.Name, XqlTableNameMap.Map(lo.Value.Name, ws.Value!.Name), hRow, hCol, dr, dc);
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
            return OnExcelThreadAsync<(string?, string?)>(() =>
            {
                using var app = SmartCom<Excel.Application>.Wrap((Excel.Application)ExcelDnaUtil.Application);
                using var rng = SmartCom<Excel.Range>.Wrap(app.Value?.Selection as Excel.Range);
                if (rng?.Value == null) return (null, null);

                using var ws = SmartCom<Excel.Worksheet>.Wrap((Excel.Worksheet)rng.Value.Worksheet);
                using var lo = SmartCom<Excel.ListObject>.Wrap(rng.Value.ListObject ?? XqlSheet.FindListObjectContaining(ws.Value!, rng.Value));
                if (lo?.Value?.HeaderRowRange == null) return (ws?.Value?.Name, null);

                string tableName = XqlTableNameMap.Map(lo.Value.Name, ws.Value!.Name);
                int hRow = lo.Value.HeaderRowRange.Row;
                int hCol = lo.Value.HeaderRowRange.Column;
                int rowOffset = rng.Value.Row - (hRow + 1);
                int colOffset = rng.Value.Column - hCol;

                return (ws.Value!.Name, XqlSheet.CellKey(ws.Value!.Name, tableName, hRow, hCol, rowOffset, colOffset));
            });
        }

        // (선택) 키를 더블클릭으로 점프할 때 사용: 상대키 → 현재 Range 복원
        public static Task<bool> TryJumpToAsync(string key, CancellationToken ct = default)
        {
            if (string.IsNullOrWhiteSpace(key))
                return Task.FromResult(false);

            var tcs = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);

            _ = XqlCommon.BridgeAsync<string, XqlSheet.RelDesc?>(
                // UI hop: 반드시 UI에서 할 필요는 없지만, Bridge의 형태를 맞춰 둠
                captureOnUi: () => key.Trim(),

                // BG hop: 파싱(순수 연산, COM 접근 금지)
                workOnBg: (k, token) =>
                {
                    var ok = XqlSheet.TryParse(k, out var desc);
                    return Task.FromResult(ok ? desc : (XqlSheet.RelDesc?)null);
                },

                // UI hop: COM 접근 및 Select
                applyOnUi: desc =>
                {
                    if (desc == null) { tcs.TrySetResult(false); return; }

                    try
                    {
                        var app = (Excel.Application)ExcelDnaUtil.Application;

                        if (XqlSheet.TryResolve(app, desc.Value, out var range, out _, out _))
                        {
                            using var _rg = SmartCom<Excel.Range>.Wrap(range);
                            _rg.Value?.Select();
                            tcs.TrySetResult(_rg.Value != null);
                        }
                        else
                        {
                            tcs.TrySetResult(false);
                        }
                    }
                    catch
                    {
                        tcs.TrySetResult(false);
                    }
                },
                ct
            )
            .ContinueWith(t =>
            {
                // Bridge 단계에서 취소/예외가 나면 결과를 정리
                if (t.IsCanceled) tcs.TrySetCanceled(ct);
                else if (t.IsFaulted) tcs.TrySetResult(false);
            }, TaskScheduler.Default);

            return tcs.Task;
        }

    }
}
