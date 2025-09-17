using System.Drawing;
using System.Windows.Forms;

namespace XQLite.AddIn
{
    public sealed class XqlInspectorForm : Form
    {
        private static XqlInspectorForm? _inst;
        private ListView lv = new();
        private Timer auto = new();
        private CheckBox chk = new() { Text = "Auto Refresh" };
        private Button btnRefresh = new() { Text = "Refresh" };
        private Button btnClear = new() { Text = "Clear" };

        public static void ShowSingleton()
        {
            if (_inst == null || _inst.IsDisposed) _inst = new XqlInspectorForm();
            _inst.Show(); _inst.BringToFront();
        }

        public XqlInspectorForm()
        {
            Text = "XQLite Inspector"; StartPosition = FormStartPosition.CenterScreen; Width = 1000; Height = 520;
            lv.View = View.Details; lv.FullRowSelect = true; lv.Dock = DockStyle.Fill;
            lv.Columns.AddRange(new[]{
            new ColumnHeader{Text="Time", Width=120},
            new ColumnHeader{Text="Level", Width=60},
            new ColumnHeader{Text="Table", Width=160},
            new ColumnHeader{Text="Message", Width=480},
            new ColumnHeader{Text="Detail", Width=160}
        });

            var top = new FlowLayoutPanel { Dock = DockStyle.Top, Height = 36, FlowDirection = FlowDirection.LeftToRight };
            top.Controls.AddRange(new Control[] { chk, btnRefresh, btnClear });
            Controls.Add(lv); Controls.Add(top);

            auto.Interval = 1000; auto.Tick += (_, __) => { if (chk.Checked) RefreshLogs(); };
            auto.Start();
            btnRefresh.Click += (_, __) => RefreshLogs();
            btnClear.Click += (_, __) => lv.Items.Clear();

            Load += (_, __) => RefreshLogs();
        }

        private void RefreshLogs()
        {
            var items = XqlFileLogger.TakeLogs(300);
            foreach (var it in items)
            {
                var detail = (it.Details != null && it.Details.Length > 0)
                    ? string.Join("; ", it.Details)
                    : string.Empty;

                var lvi = new ListViewItem(new[]
                {
            it.At.ToString("HH:mm:ss"),
            it.Level,
            it.Table,
            it.Message,
            detail
        });
                lv.Items.Add(lvi);
            }
            if (lv.Items.Count > 0)
                lv.EnsureVisible(lv.Items.Count - 1);
        }
    }
}