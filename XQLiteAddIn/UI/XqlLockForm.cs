using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using ExcelDna.Integration;
using Excel = Microsoft.Office.Interop.Excel;

namespace XQLite.AddIn
{
    public sealed class XqlLockForm : Form
    {
        private static XqlLockForm? _inst;
        internal static void ShowSingleton()
        {
            if (_inst == null || _inst.IsDisposed) _inst = new XqlLockForm();
            _inst.Show();
            _inst.BringToFront();
        }

        // 단일 Collab 인스턴스(프로젝트 전역에서 사용 중인 것)
        private XqlCollab Collab => XqlAddIn.Collab!; // 필요 시 참조 수정

        private readonly Button btnLockCol = new() { Text = "Lock Column", Width = 120, Dock = DockStyle.Left };
        private readonly Button btnLockCell = new() { Text = "Lock Cell", Width = 100, Dock = DockStyle.Left };
        private readonly Button btnUnlockAll = new() { Text = "Unlock (Mine)", Width = 120, Dock = DockStyle.Left };
        private readonly Button btnJump = new() { Text = "Jump to Selected", Width = 140, Dock = DockStyle.Left };

        private readonly ListView lv = new() { View = View.Details, Dock = DockStyle.Fill, FullRowSelect = true };
        private readonly StatusStrip status = new();
        private readonly ToolStripStatusLabel lbl = new() { Text = "Ready" };

        // 서버 목록 API가 없으므로 로컬 히스토리로 최근 락을 보관
        private readonly LinkedList<string> _recentKeys = new();
        private const int MaxRecent = 200;

        public XqlLockForm()
        {
            Text = "XQLite Locks";
            StartPosition = FormStartPosition.Manual;
            Left = 80; Top = 80; Width = 760; Height = 420;

            var top = new FlowLayoutPanel { Dock = DockStyle.Top, Height = 40, Padding = new Padding(8), AutoSize = false };
            top.Controls.AddRange(new Control[] { btnLockCol, btnLockCell, btnUnlockAll, btnJump });

            lv.Columns.AddRange(new[]
            {
                new ColumnHeader { Text = "Key (relative)", Width = 440 },
                new ColumnHeader { Text = "When",           Width = 180 },
                new ColumnHeader { Text = "Note",           Width = 100 },
            });

            status.Items.Add(lbl);

            Controls.Add(lv);
            Controls.Add(top);
            Controls.Add(status);

            // 핸들러
            btnLockCol.Click += async (_, __) => await LockCurrentColumn();
            btnLockCell.Click += async (_, __) => await LockCurrentCell();
            btnUnlockAll.Click += async (_, __) => await UnlockMine();
            btnJump.Click += (_, __) => JumpToSelected();

            lv.ItemActivate += (_, __) => JumpToSelected();

            Load += (_, __) => RefreshList();
        }

        // ─────────────────────────────────────────────────────────────────────
        // Actions
        // ─────────────────────────────────────────────────────────────────────

        private async Task LockCurrentColumn()
        {
            try
            {
                if (Collab == null) { Warn("Collab not ready."); return; }

                // 선택된 컬럼으로 상대키 생성 → Acquire (Collab 내부에서 마이그레이션 처리)
                if (await Collab.AcquireCurrentColumn().ConfigureAwait(false))
                {
                    var key = TryCurrentColumnKeyOrNull();
                    if (!string.IsNullOrWhiteSpace(key)) PushRecent(key!, "locked");
                    Info("Column locked.");
                }
                else Warn("Lock failed.");
            }
            catch (Exception ex) { Warn("Lock error: " + ex.Message); }
            RefreshList();
        }

        private async Task LockCurrentCell()
        {
            try
            {
                if (Collab == null) { Warn("Collab not ready."); return; }

                if (await Collab.AcquireCurrentCell().ConfigureAwait(false))
                {
                    var key = TryCurrentCellKeyOrNull();
                    if (!string.IsNullOrWhiteSpace(key)) PushRecent(key!, "locked");
                    Info("Cell locked.");
                }
                else Warn("Lock failed.");
            }
            catch (Exception ex) { Warn("Lock error: " + ex.Message); }
            RefreshList();
        }

        private async Task UnlockMine()
        {
            try
            {
                if (Collab == null) { Warn("Collab not ready."); return; }
                if (await Collab.ReleaseByMe().ConfigureAwait(false))
                {
                    Info("Released my locks.");
                }
                else Warn("Release failed.");
            }
            catch (Exception ex) { Warn("Release error: " + ex.Message); }
        }

        private void JumpToSelected()
        {
            try
            {
                if (lv.SelectedItems.Count == 0) return;
                var key = lv.SelectedItems[0].SubItems[0].Text;
                if (string.IsNullOrWhiteSpace(key)) return;

                if (XqlCollab.TryJumpTo(key))
                {
                    Info("Jumped.");
                }
                else Warn("Cannot resolve key.");
            }
            catch (Exception ex) { Warn("Jump error: " + ex.Message); }
        }

        // ─────────────────────────────────────────────────────────────────────
        // Helpers
        // ─────────────────────────────────────────────────────────────────────

        private void PushRecent(string key, string note)
        {
            // 중복은 맨 앞으로 당김
            var node = _recentKeys.FirstOrDefault(k => string.Equals(k, key, StringComparison.Ordinal));
            if (node != null)
            {
                _recentKeys.Remove(key);
            }
            _recentKeys.AddFirst(key);
            while (_recentKeys.Count > MaxRecent) _recentKeys.RemoveLast();

            // UI 반영
            var li = new ListViewItem(new[] { key, DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"), note });
            lv.Items.Insert(0, li);
            // 리스트가 너무 커졌으면 정리
            while (lv.Items.Count > MaxRecent) lv.Items.RemoveAt(lv.Items.Count - 1);
        }

        private void RefreshList()
        {
            // 첫 로드 시 아무 것도 없으면 현재 선택으로 힌트 제공
            if (lv.Items.Count == 0)
            {
                var hint = TryCurrentCellKeyOrNull() ?? TryCurrentColumnKeyOrNull();
                if (!string.IsNullOrWhiteSpace(hint))
                    PushRecent(hint!, "hint");
            }
        }

        private static string? TryCurrentCellKeyOrNull()
        {
            try
            {
                var app = (Excel.Application)ExcelDnaUtil.Application;
                var rng = (Excel.Range)app.Selection;
                if (rng == null) return null;

                var ws = (Excel.Worksheet)rng.Worksheet;
                var lo = rng.ListObject ?? XqlSheetUtil.FindListObjectContaining(ws, rng);
                if (lo?.HeaderRowRange == null) return null;

                var tableName = XqlTableNameMap.Map(lo.Name, ws.Name);
                int hRow = lo.HeaderRowRange.Row;
                int hCol = lo.HeaderRowRange.Column;
                int rowOffset = rng.Row - (hRow + 1);
                int colOffset = rng.Column - hCol;

                return $"cell:{ws.Name}:{tableName}:H{hRow}C{hCol}:dr={rowOffset}:dc={colOffset}";
            }
            catch { return null; }
        }

        private static string? TryCurrentColumnKeyOrNull()
        {
            try
            {
                var app = (Excel.Application)ExcelDnaUtil.Application;
                var rng = (Excel.Range)app.Selection;
                if (rng == null) return null;

                var ws = (Excel.Worksheet)rng.Worksheet;
                var lo = rng.ListObject ?? XqlSheetUtil.FindListObjectContaining(ws, rng);
                if (lo?.HeaderRowRange == null) return null;

                var (header, headers) = XqlSheetUtil.GetHeaderAndNames(ws);
                int colIndex = rng.Column - lo.HeaderRowRange.Column; // 0-base
                if (colIndex < 0 || colIndex >= headers.Count) return null;

                string headerName = headers[colIndex];
                var tableName = XqlTableNameMap.Map(lo.Name, ws.Name);
                int hRow = lo.HeaderRowRange.Row;
                int hCol = lo.HeaderRowRange.Column;
                int colOffset = colIndex;

                return $"col:{ws.Name}:{tableName}:H{hRow}C{hCol}:dx={colOffset}:hdr={Escape(headerName)}";
            }
            catch { return null; }
        }

        private static string Escape(string s) => s.Replace("\\", "\\\\").Replace(":", "\\:");

        private void Info(string msg) { lbl.Text = msg; }
        private void Warn(string msg) { lbl.Text = "⚠ " + msg; }
    }
}
