using ExcelDna.Integration;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
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

        private readonly Button btnLockCol = new() { Text = "Lock Column", Width = 120, Dock = DockStyle.Left };
        private readonly Button btnLockCell = new() { Text = "Lock Cell", Width = 100, Dock = DockStyle.Left };
        private readonly Button btnUnlock = new() { Text = "Unlock (Mine)", Width = 140, Dock = DockStyle.Left };
        private readonly ListView lv = new() { View = View.Details, Dock = DockStyle.Fill, FullRowSelect = true };
        private readonly Timer auto = new() { Interval = 2000 };

        public XqlLockForm()
        {
            Text = "XQLite Locks";
            StartPosition = FormStartPosition.Manual;
            Left = 60; Top = 60; Width = 720; Height = 360;

            var top = new FlowLayoutPanel { Dock = DockStyle.Top, Height = 36 };
            top.Controls.AddRange(new Control[] { btnLockCol, btnLockCell, btnUnlock });

            lv.Columns.AddRange(new[]
            {
                new ColumnHeader { Text = "LockId",   Width = 220 }, // 리소스 키
                new ColumnHeader { Text = "Owner",    Width = 120 },
                new ColumnHeader { Text = "Resource", Width = 220 },
                new ColumnHeader { Text = "Expires",  Width = 120 }
            });

            Controls.Add(lv);
            Controls.Add(top);

            btnLockCol.Click += async (_, __) => await LockCurrentColumnAsync();
            btnLockCell.Click += async (_, __) => await LockCurrentCellAsync();
            btnUnlock.Click += async (_, __) => await UnlockMineAsync();

            auto.Tick += (_, __) => RefreshList();
            auto.Start();
            Load += (_, __) => RefreshList();
        }

        // ===== Actions =====

        private async Task LockCurrentColumnAsync()
        {
            var sel = GetSelection();
            if (sel == null || string.IsNullOrEmpty(sel.Table) || string.IsNullOrEmpty(sel.ColumnName))
            {
                MessageBox.Show("선택 영역이 테이블 컬럼이 아닙니다.");
                return;
            }
            if (XqlCollab.Instance == null)
            {
                MessageBox.Show("Collab 모듈이 초기화되지 않았습니다.");
                return;
            }

            var nickname = XqlAddIn.Cfg?.Nickname ?? "anonymous";
#pragma warning disable CS8602 // null 가능 참조에 대한 역참조입니다.
            var ok = await XqlAddIn.Collab.TryAcquireColumnLock(sel.Table!, sel.ColumnName!, nickname);
#pragma warning restore CS8602 // null 가능 참조에 대한 역참조입니다.
            MessageBox.Show(ok ? "Locked" : "Lock failed");
            RefreshList();
        }

        private async Task LockCurrentCellAsync()
        {
            var sel = GetSelection();
            if (sel == null)
                return;
            if (XqlCollab.Instance == null)
            {
                MessageBox.Show("Collab 모듈이 초기화되지 않았습니다.");
                return;
            }

            var nickname = XqlAddIn.Cfg?.Nickname ?? "anonymous";
            var ok = await XqlCollab.Instance.TryAcquireCellLock(sel.Sheet, sel.Address, nickname);
            MessageBox.Show(ok ? "Locked" : "Lock failed");
            RefreshList();
        }

        private async Task UnlockMineAsync()
        {
            if (XqlCollab.Instance == null)
            {
                MessageBox.Show("Collab 모듈이 초기화되지 않았습니다.");
                return;
            }
            var nick = XqlAddIn.Cfg?.Nickname ?? "anonymous";
            XqlCollab.Instance.ReleaseLocksBy(nick); // 서버/로컬 모두 해제
            await Task.Delay(50);
            RefreshList();
        }

        // ===== List Rendering =====

        private void RefreshList()
        {
            try
            {
                lv.BeginUpdate();
                lv.Items.Clear();

                foreach (var it in SnapshotLocks())
                {
                    var li = new ListViewItem(new[]
                    {
                        it.LockId,
                        it.Owner,
                        it.Resource,
                        it.ExpiresAtLocal
                    })
                    {
                        Tag = it.LockId
                    };
                    lv.Items.Add(li);
                }
            }
            finally
            {
                lv.EndUpdate();
            }
        }

        // 우선 Collab에 공개 API가 있다면 사용, 없으면 reflection fallback
        private IEnumerable<LockView> SnapshotLocks()
        {
            var list = new List<LockView>();

            if (XqlCollab.Instance != null)
            {
                // 1) 공개 API가 있다면 사용
                var pub = XqlCollab.Instance.GetType().GetMethod("GetCurrentLocks", BindingFlags.Instance | BindingFlags.Public);
                if (pub != null)
                {
                    var result = pub.Invoke(XqlCollab.Instance, null) as IEnumerable<object>;
                    if (result != null)
                    {
                        foreach (var o in result)
                        {
                            // 익명/튜플 등에 대응
                            var t = o.GetType();
                            string key = TryProp<string>(t, o, "Key") ?? TryProp<string>(t, o, "Resource") ?? "";
                            string owner = TryProp<string>(t, o, "Owner") ?? TryProp<string>(t, o, "By") ?? "";
                            DateTime? exp = TryProp<DateTime?>(t, o, "ExpiresAt");
                            list.Add(new LockView(key, owner, key, ToLocalString(exp)));
                        }
                        return list;
                    }
                }

                // 2) 내부 딕셔너리(_locks: ConcurrentDictionary<string,string>)에 대한 reflection
                var fld = XqlCollab.Instance?.GetType().GetField("_locks", BindingFlags.NonPublic | BindingFlags.Instance);
                if (fld?.GetValue(XqlCollab.Instance) is ConcurrentDictionary<string, string> dict)
                {
                    foreach (var kv in dict)
                        list.Add(new LockView(kv.Key, kv.Value, kv.Key, "")); // TTL은 서버 기준이라 미표시
                }
            }

            return list;
        }

        private static T? TryProp<T>(Type t, object o, string name)
        {
            var p = t.GetProperty(name, BindingFlags.Public | BindingFlags.Instance);
            if (p == null) return default;
            try { return (T?)p.GetValue(o); } catch { return default; }
        }

        private static string ToLocalString(DateTime? dt)
        {
            if (!dt.HasValue) return "";
            try { return dt.Value.ToLocalTime().ToString("yyyy-MM-dd HH:mm:ss"); }
            catch { return ""; }
        }

        // ===== Selection Helper =====

        private sealed class SelInfo
        {
            public string Sheet = "";
            public string Address = "";
            public string? Table;
            public string? ColumnName;
        }

        private SelInfo? GetSelection()
        {
            try
            {
                var app = (Excel.Application)ExcelDnaUtil.Application;
                var rng = (Excel.Range)app.Selection;
                if (rng == null) return null;

                var ws = (Excel.Worksheet)rng.Worksheet;
                string sheet = ws.Name;
                string addr = rng.Address[false, false];

                string? tableName = null;
                string? colName = null;

                if (rng.ListObject is Excel.ListObject lo)
                {
                    tableName = XqlTableNameMap.Map(lo.Name, ws.Name);
                    int colIndex = rng.Column - lo.HeaderRowRange.Column + 1;
                    if (colIndex >= 1 && colIndex <= lo.HeaderRowRange.Columns.Count)
                    {
                        var headerArr = (object[,])lo.HeaderRowRange.Value2;
                        colName = Convert.ToString(headerArr[1, colIndex]) ?? "";
                    }
                }

                return new SelInfo
                {
                    Sheet = sheet,
                    Address = addr,
                    Table = tableName,
                    ColumnName = colName
                };
            }
            catch
            {
                return null;
            }
        }

        // 목록 표시용 뷰모델
        private readonly record struct LockView(string LockId, string Owner, string Resource, string ExpiresAtLocal);
    }
}
