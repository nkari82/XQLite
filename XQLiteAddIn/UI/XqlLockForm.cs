using ExcelDna.Integration;
using System;
using System.Threading.Tasks;
using System.Windows.Forms;
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

        private Button btnLockCol = new() { Text = "Lock Column", Width = 120, Dock = DockStyle.Left };
        private Button btnLockCell = new() { Text = "Lock Cell", Width = 100, Dock = DockStyle.Left };
        private Button btnUnlock = new() { Text = "Unlock Selected", Width = 140, Dock = DockStyle.Left };
        private ListView lv = new() { View = View.Details, Dock = DockStyle.Fill, FullRowSelect = true };
        private Timer auto = new() { Interval = 2000 };

        public XqlLockForm()
        {
            Text = "XQLite Locks"; StartPosition = FormStartPosition.Manual; Left = 60; Top = 60; Width = 720; Height = 360;
            var top = new FlowLayoutPanel { Dock = DockStyle.Top, Height = 36 }; top.Controls.AddRange(new Control[] { btnLockCol, btnLockCell, btnUnlock });
            lv.Columns.AddRange(new[] { new ColumnHeader { Text = "LockId", Width = 180 }, new ColumnHeader { Text = "Owner", Width = 120 }, new ColumnHeader { Text = "Resource", Width = 220 }, new ColumnHeader { Text = "Expires", Width = 160 } });
            Controls.Add(lv); Controls.Add(top);
            btnLockCol.Click += async (_, __) => await LockCurrentColumnAsync();
            btnLockCell.Click += async (_, __) => await LockCurrentCellAsync();
            btnUnlock.Click += async (_, __) => await UnlockSelectedAsync();
            auto.Tick += (_, __) => RefreshList(); auto.Start(); Load += (_, __) => RefreshList();
        }

        private async Task LockCurrentColumnAsync()
        {
            var sel = GetSelection(); if (sel == null) return;
            if (string.IsNullOrEmpty(sel.Table) || string.IsNullOrEmpty(sel.ColumnName)) { MessageBox.Show("선택 영역이 테이블 컬럼이 아닙니다."); return; }

            if (sel.Table is null || sel.ColumnName is null)
            {
                MessageBox.Show("Invalid selection: table or column is null");
                return;
            }

            var ok = await XqlLockService.AcquireColumnAsync(sel.Table, sel.ColumnName);
            MessageBox.Show(ok ? "Locked" : "Lock failed");
            RefreshList();
        }

        private async Task LockCurrentCellAsync()
        {
            var sel = GetSelection(); if (sel == null) return;
            var ok = await XqlLockService.AcquireCellAsync(sel.Sheet, sel.Address);
            MessageBox.Show(ok ? "Locked" : "Lock failed");
            RefreshList();
        }

        private async Task UnlockSelectedAsync()
        {
            if (lv.SelectedItems.Count == 0) return; 
            var id = lv.SelectedItems[0].Tag as string; 
            if (id == null || id == "") 
                return;

            await XqlLockService.ReleaseAsync(id); RefreshList();
        }

        private void RefreshList()
        {
            // 내부 캐시 반영 (서비스가 주기적으로 갱신)
            var f = typeof(XqlLockService).GetField("_locksById", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
            var dict = f?.GetValue(null) as System.Collections.IDictionary; if (dict == null) return;
            lv.BeginUpdate(); lv.Items.Clear();
            foreach (System.Collections.DictionaryEntry de in dict)
            {
                var li = (XqlLockService.LockInfo)de.Value!;
                lv.Items.Add(new ListViewItem(new[] { li.lockId, li.owner, li.resource, li.expiresAt.ToLocalTime().ToString("yyyy-MM-dd HH:mm:ss") }) { Tag = li.lockId });
            }
            lv.EndUpdate();
        }

        private sealed class SelInfo { public string Sheet = ""; public string Address = ""; public string? Table; public string? ColumnName; }
        private SelInfo? GetSelection()
        {
            try
            {
                var app = (Excel.Application)ExcelDnaUtil.Application; var rng = (Excel.Range)app.Selection; if (rng == null) return null;
                var ws = (Excel.Worksheet)rng.Worksheet; string sheet = ws.Name; string addr = rng.Address[false, false];
                string? tableName = null; string? colName = null;
                if (rng.ListObject is Excel.ListObject lo)
                {
                    tableName = XqlTableNameMap.Map(lo.Name, ws.Name);
                    int colIndex = rng.Column - lo.HeaderRowRange.Column + 1;
                    if (colIndex >= 1 && colIndex <= lo.HeaderRowRange.Columns.Count)
                    {
                        var headerArr = (object[,])lo.HeaderRowRange.Value2; colName = Convert.ToString(headerArr[1, colIndex]) ?? "";
                    }
                }
                return new SelInfo { Sheet = sheet, Address = addr, Table = tableName, ColumnName = colName };
            }
            catch { return null; }
        }
    }
}