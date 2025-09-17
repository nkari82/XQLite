using System;
using System.Drawing;
using System.Windows.Forms;

namespace XQLite.AddIn
{
    public sealed class XqlConfigForm : Form
    {
        private static XqlConfigForm? _instance;

        private TextBox txtEndpoint = new();
        private TextBox txtApiKey = new();
        private TextBox txtNickname = new();
        private TextBox txtProject = new();
        private NumericUpDown numPull = new();
        private NumericUpDown numDebounce = new();
        private NumericUpDown numHeartbeat = new();
        private NumericUpDown numLockTtl = new();
        private Button btnSaveApply = new();
        private Button btnSave = new();
        private Button btnClose = new();

        private CheckBox chkSecure = new() { Text = "Protect API Key (DPAPI)", AutoSize = true };

        internal static void ShowSingleton()
        {
            if (_instance == null || _instance.IsDisposed)
                _instance = new XqlConfigForm();
            _instance.Show();
            _instance.BringToFront();
        }

        public XqlConfigForm()
        {
            Text = "XQLite Config";
            StartPosition = FormStartPosition.CenterScreen;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false; MinimizeBox = false;
            AutoScaleMode = AutoScaleMode.Font; AutoSize = true; AutoSizeMode = AutoSizeMode.GrowAndShrink;

            var layout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, Padding = new Padding(12), AutoSize = true };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));

            AddRow(layout, "Endpoint", txtEndpoint);
            AddRow(layout, "API Key", txtApiKey, password: true);
            AddRow(layout, "Nickname", txtNickname);
            AddRow(layout, "Project", txtProject);
            AddRow(layout, "Pull (sec)", numPull, 1, 3600, 10);
            AddRow(layout, "Debounce (ms)", numDebounce, 100, 10000, 2000);
            AddRow(layout, "Heartbeat (sec)", numHeartbeat, 1, 120, 3);
            AddRow(layout, "Lock TTL (sec)", numLockTtl, 1, 600, 10);

            AddSecurityRow(layout);

            var buttons = new FlowLayoutPanel { FlowDirection = FlowDirection.RightToLeft, Dock = DockStyle.Fill, AutoSize = true };
            btnClose.Text = "Close"; btnClose.Click += (_, __) => Close();
            btnSave.Text = "Save"; btnSave.Click += (_, __) => SaveConfig(apply: false);
            btnSaveApply.Text = "Save && Apply"; btnSaveApply.Click += (_, __) => SaveConfig(apply: true);
            buttons.Controls.AddRange(new Control[] { btnClose, btnSave, btnSaveApply });
            layout.Controls.Add(buttons, 0, layout.RowCount); layout.SetColumnSpan(buttons, 2);

            Controls.Add(layout);
            Load += (_, __) => LoadConfigToUi();
        }

        private void AddRow(TableLayoutPanel p, string label, Control editor, int min = 0, int max = 0, int val = 0, bool password = false)
        {
            var l = new Label { Text = label, AutoSize = true, TextAlign = ContentAlignment.MiddleLeft, Padding = new Padding(0, 6, 8, 0) };
            p.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            p.Controls.Add(l, 0, p.RowCount);

            if (editor is TextBox tb)
            {
                tb.Width = 420; if (password) tb.UseSystemPasswordChar = true;
            }
            if (editor is NumericUpDown nud)
            {
                nud.Minimum = min; nud.Maximum = max; nud.Value = val; nud.Width = 120;
            }
            p.Controls.Add(editor, 1, p.RowCount - 1);
        }

        private void AddSecurityRow(TableLayoutPanel layout)
        {
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.Controls.Add(new Label { Text = "Security", AutoSize = true, Padding = new Padding(0, 6, 8, 0) }, 0, layout.RowCount);
            layout.Controls.Add(chkSecure, 1, layout.RowCount);
        }

        private void LoadConfigToUi()
        {
            var cfg = XqlAddIn.Cfg ?? XqlConfig.Load();
            txtEndpoint.Text = cfg.Endpoint;
            txtApiKey.Text = cfg.ApiKey;
            txtNickname.Text = cfg.Nickname;
            txtProject.Text = cfg.Project;
            numPull.Value = Clamp(cfg.PullSec, (int)numPull.Minimum, (int)numPull.Maximum);
            numDebounce.Value = Clamp(cfg.DebounceMs, (int)numDebounce.Minimum, (int)numDebounce.Maximum);
            numHeartbeat.Value = Clamp(cfg.HeartbeatSec, (int)numHeartbeat.Minimum, (int)numHeartbeat.Maximum);
            numLockTtl.Value = Clamp(cfg.LockTtlSec, (int)numLockTtl.Minimum, (int)numLockTtl.Maximum);
        }

        private void SaveConfig(bool apply)
        {
            var cfg = XqlAddIn.Cfg ?? new XqlConfig();
            cfg.Endpoint = txtEndpoint.Text.Trim();
            cfg.ApiKey = chkSecure.Checked ? "__DPAPI__" : txtApiKey.Text; // 표시용 플래그
            cfg.Nickname = txtNickname.Text.Trim();
            cfg.Project = txtProject.Text.Trim();
            cfg.PullSec = (int)numPull.Value;
            cfg.DebounceMs = (int)numDebounce.Value;
            cfg.HeartbeatSec = (int)numHeartbeat.Value;
            cfg.LockTtlSec = (int)numLockTtl.Value;

            cfg.Save(preferSidecar: true);
            if (chkSecure.Checked) XqlSecrets.SaveApiKey(txtApiKey.Text); else XqlSecrets.Clear();
            XqlAddIn.Cfg = cfg;
            if (apply) XqlAddIn.RestartRuntime(cfg);
            MessageBox.Show(this, apply ? "Saved & Applied." : "Saved.", "XQLite", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void InitializeComponent()
        {

        }

        private static int Clamp(int v, int lo, int hi) => Math.Max(lo, Math.Min(v, hi));
    }
}