using System;
using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;

namespace XQLite.AddIn
{
    public sealed class XqlConfigForm : Form
    {
        private static XqlConfigForm? _instance;

        // ── width presets (px)
        const int W_LABEL = 150;
        const int W_NUM = 90;
        const int W_EP = 360; // Endpoint (wide)
        const int W_KEY = 160; // API Key (short)
        const int W_PROJ = 240; // Project
        const int W_NICK = 180; // Nickname

        // ── Fields
        private readonly TextBox txtEndpoint = new();
        private readonly TextBox txtApiKey = new();
        private readonly TextBox txtNickname = new();
        private readonly TextBox txtProject = new();
        private readonly NumericUpDown numPull = new();
        private readonly NumericUpDown numDebounce = new();
        private readonly NumericUpDown numHeartbeat = new();
        private readonly NumericUpDown numLockTtl = new();

        private readonly Button btnSaveApply = new();
        private readonly Button btnSave = new();
        private readonly Button btnClose = new();
        private readonly CheckBox chkSecure = new() { Text = "Protect API Key (DPAPI)", AutoSize = true };

        private readonly ToolTip tips = new() { AutoPopDelay = 8000, InitialDelay = 300, ReshowDelay = 100, UseFading = true, UseAnimation = true };
        private readonly ErrorProvider errors = new() { BlinkStyle = ErrorBlinkStyle.NeverBlink };
        private bool _apiKeyMasked = true;

        internal static void ShowSingleton()
        {
            if (_instance == null || _instance.IsDisposed)
                _instance = new XqlConfigForm();
            _instance.Show();
            _instance.BringToFront();
        }

        public XqlConfigForm()
        {
            // ── Form
            Text = "XQLite Settings";
            StartPosition = FormStartPosition.CenterScreen;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false; MinimizeBox = false;
            AutoScaleMode = AutoScaleMode.Dpi;
            ClientSize = new Size(520, 460); // 타이트하게
            Font = SystemFonts.MessageBoxFont;

            AcceptButton = btnSaveApply;
            CancelButton = btnClose;

            // 루트(그룹을 위에서부터 쌓기)
            var root = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                AutoSize = false,
                ColumnCount = 1,
                Padding = new Padding(12)
            };
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));

            root.Controls.Add(BuildHeader());
            root.Controls.Add(BuildGroup_Connection());
            root.Controls.Add(BuildGroup_Identity());
            root.Controls.Add(BuildGroup_Timing());
            root.Controls.Add(BuildButtons());

            Controls.Add(root);
            Load += (_, __) => LoadConfigToUi();
        }

        // ─────────────────────────────────────────────────────────────
        private Control BuildHeader()
        {
            var panel = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                ColumnCount = 2,
                Padding = new Padding(0, 0, 0, 8)
            };
            panel.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            panel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));

            var icon = new PictureBox
            {
                Image = SystemIcons.Information.ToBitmap(),
                SizeMode = PictureBoxSizeMode.CenterImage,
                Width = 40,
                Height = 40,
                Margin = new Padding(0, 0, 8, 0)
            };

            var title = new Label { Text = "XQLite Configuration", Font = new Font(Font, FontStyle.Bold), AutoSize = true };
            var subtitle = new Label
            {
                Text = "Excel ↔ SQLite 동기화 / Presence / 검증 주기 설정",
                AutoSize = true,
                ForeColor = SystemColors.GrayText,
                Margin = new Padding(0, 2, 0, 0)
            };

            var textPanel = new TableLayoutPanel { AutoSize = true, ColumnCount = 1, Dock = DockStyle.Fill };
            textPanel.Controls.Add(title, 0, 0);
            textPanel.Controls.Add(subtitle, 0, 1);

            panel.Controls.Add(icon, 0, 0);
            panel.Controls.Add(textPanel, 1, 0);
            return panel;
        }

        // ─────────────────────────────────────────────────────────────
        // Common helpers
        private static GroupBox MakeGroup(string title) => new GroupBox
        {
            Text = title,
            Dock = DockStyle.Top,
            AutoSize = true,
            Padding = new Padding(10),
            Margin = new Padding(0, 6, 0, 0)
        };

        private static TableLayoutPanel MakeGrid(int columns = 2)
        {
            var grid = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                ColumnCount = columns,
                Margin = new Padding(6, 4, 6, 0)
            };
            return grid;
        }

        private void AddRow(TableLayoutPanel grid, string label, Control editor, int editorWidth, string? tooltip = null)
        {
            int row = grid.RowCount;
            grid.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            var l = new Label
            {
                Text = label,
                AutoSize = true,
                TextAlign = ContentAlignment.MiddleLeft,
                Margin = new Padding(0, 4, 6, 4),
                Width = W_LABEL
            };
            grid.Controls.Add(l, 0, row);

            if (editor is TextBox tb)
            {
                tb.Width = editorWidth;
                tb.Anchor = AnchorStyles.Left;
            }
            else if (editor is NumericUpDown nud)
            {
                nud.Width = editorWidth;
                nud.Anchor = AnchorStyles.Left;
            }
            else if (editor is Panel p)
            {
                p.AutoSize = true;
                p.Anchor = AnchorStyles.Left;
            }

            grid.Controls.Add(editor, 1, row);
            if (!string.IsNullOrWhiteSpace(tooltip)) tips.SetToolTip(editor, tooltip);
        }

        // ─────────────────────────────────────────────────────────────
        // Connection
        private GroupBox BuildGroup_Connection()
        {
            var g = MakeGroup("Connection");

            // 2열 그리드 (Label | Editor)
            var grid = MakeGrid(2);
            grid.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, W_LABEL));
            grid.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));

            // ── API Key: 텍스트필드 "내부 우측" Show/Hide 버튼 (Interop 불필요)
            const int BtnW = 40;
            int fieldW = W_KEY + BtnW; // 버튼 포함 필드 총 너비

            var apiInline = new InlineButtonTextBox("Show", BtnW)
            {
                Width = fieldW
            };

            // 내부 TextBox 설정
            apiInline.TextBox.UseSystemPasswordChar = true;

            // 기존 txtApiKey 필드와 양방향 동기화 (Load/Save 코드 그대로 사용 가능)
            txtApiKey.UseSystemPasswordChar = true;          // 논리 상태 유지
            txtApiKey.Visible = false;                        // UI에 직접 붙이지 않음
            apiInline.TextBox.DataBindings.Add("Text", txtApiKey, "Text", false, DataSourceUpdateMode.OnPropertyChanged);
            txtApiKey.DataBindings.Add("Text", apiInline.TextBox, "Text", false, DataSourceUpdateMode.OnPropertyChanged);

            // 토글 버튼
            apiInline.TailButton.Click += (_, __) =>
            {
                apiInline.TextBox.UseSystemPasswordChar = !apiInline.TextBox.UseSystemPasswordChar;
                _apiKeyMasked = apiInline.TextBox.UseSystemPasswordChar;
                apiInline.TailButton.Text = _apiKeyMasked ? "Show" : "Hide";
            };
            // 초기 버튼 텍스트 동기화
            _apiKeyMasked = apiInline.TextBox.UseSystemPasswordChar;
            apiInline.TailButton.Text = _apiKeyMasked ? "Show" : "Hide";

            AddRow(grid, "API &Key", apiInline, fieldW);

            // ── Security
            AddRow(grid, "Security", chkSecure, 0, "API Key를 DPAPI로 암호화 저장");

            // ── Endpoint (넓게)
            txtEndpoint.Width = W_EP;
            AddRow(grid, "&Endpoint", txtEndpoint, W_EP, "예: http://localhost:48080/graphql");

            g.Controls.Add(grid);
            return g;
        }

        // ─────────────────────────────────────────────────────────────
        // Identity
        private GroupBox BuildGroup_Identity()
        {
            var g = MakeGroup("Identity");
            var grid = MakeGrid(2);
            grid.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, W_LABEL));
            grid.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));

            txtProject.Width = W_PROJ;
            txtNickname.Width = W_NICK;
            AddRow(grid, "Project", txtProject, W_PROJ);
            AddRow(grid, "Nickname", txtNickname, W_NICK);

            g.Controls.Add(grid);
            return g;
        }

        // ─────────────────────────────────────────────────────────────
        // Timing: compact 2×2 matrix  (Label+Nud | Label+Nud)
        private GroupBox BuildGroup_Timing()
        {
            var g = MakeGroup("Timing  Sync");
            var grid = MakeGrid(4);
            grid.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, W_LABEL)); // L1
            grid.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, W_NUM));   // N1
            grid.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, W_LABEL)); // L2
            grid.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, W_NUM));   // N2

            void addPair(string l1, NumericUpDown n1, string l2, NumericUpDown n2)
            {
                int row = grid.RowCount;
                grid.RowStyles.Add(new RowStyle(SizeType.AutoSize));

                var a = new Label { Text = l1, AutoSize = true, Margin = new Padding(0, 4, 6, 4), Width = W_LABEL };
                var b = new Label { Text = l2, AutoSize = true, Margin = new Padding(8, 4, 6, 4), Width = W_LABEL };
                n1.Width = W_NUM; n1.Anchor = AnchorStyles.Left;
                n2.Width = W_NUM; n2.Anchor = AnchorStyles.Left;

                grid.Controls.Add(a, 0, row);
                grid.Controls.Add(n1, 1, row);
                grid.Controls.Add(b, 2, row);
                grid.Controls.Add(n2, 3, row);
            }

            SetupNud(numLockTtl, 1, 600, 10, 1);
            SetupNud(numHeartbeat, 1, 120, 3, 1);
            SetupNud(numDebounce, 100, 10000, 2000, 50);
            SetupNud(numPull, 1, 3600, 10, 1);

            addPair("LockTTL(sec)", numLockTtl, "Heartbeat(sec)", numHeartbeat);
            addPair("Debounce(ms)", numDebounce, "Pull(sec)", numPull);

            g.Controls.Add(grid);
            return g;
        }

        private static void SetupNud(NumericUpDown nud, int min, int max, int val, int step)
        {
            nud.Minimum = min; nud.Maximum = max; nud.Value = val;
            nud.Increment = step; nud.ThousandsSeparator = true;
            nud.Margin = new Padding(0, 0, 0, 0);
        }

        // ─────────────────────────────────────────────────────────────
        private Control BuildButtons()
        {
            var panel = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.RightToLeft,
                Dock = DockStyle.Top,
                AutoSize = true,
                Padding = new Padding(0, 10, 0, 0)
            };

            btnClose.Text = "Close"; btnClose.AutoSize = true; btnClose.Click += (_, __) => Close();
            btnSave.Text = "Save"; btnSave.AutoSize = true; btnSave.Click += (_, __) => { if (ValidateChildren()) SaveConfig(apply: false); };
            btnSaveApply.Text = "Apply"; btnSaveApply.AutoSize = true; btnSaveApply.Click += (_, __) => { if (ValidateChildren()) SaveConfig(apply: true); };

            panel.Controls.AddRange(new Control[] { btnClose, btnSave, btnSaveApply });
            return panel;
        }

        // ─────────────────────────────────────────────────────────────
        private void LoadConfigToUi()
        {
            XqlConfig.Load();

            txtEndpoint.Text = XqlConfig.Endpoint;
            txtApiKey.Text = XqlConfig.ApiKey;
            txtNickname.Text = XqlConfig.Nickname;
            txtProject.Text = XqlConfig.Project;

            numPull.Value = Clamp(XqlConfig.PullSec, (int)numPull.Minimum, (int)numPull.Maximum);
            numDebounce.Value = Clamp(XqlConfig.DebounceMs, (int)numDebounce.Minimum, (int)numDebounce.Maximum);
            numHeartbeat.Value = Clamp(XqlConfig.HeartbeatSec, (int)numHeartbeat.Minimum, (int)numHeartbeat.Maximum);
            numLockTtl.Value = Clamp(XqlConfig.LockTtlSec, (int)numLockTtl.Minimum, (int)numLockTtl.Maximum);

            if (string.Equals(XqlConfig.ApiKey, "__DPAPI__", StringComparison.Ordinal))
            {
                chkSecure.Checked = true;
                txtApiKey.PlaceholderTextSafe("Stored securely (DPAPI)");
                txtApiKey.Clear();
            }
        }

        private void SaveConfig(bool apply)
        {
            XqlConfig.Endpoint = txtEndpoint.Text.Trim();
            XqlConfig.ApiKey = chkSecure.Checked ? "__DPAPI__" : txtApiKey.Text;
            XqlConfig.Nickname = txtNickname.Text.Trim();
            XqlConfig.Project = txtProject.Text.Trim();
            XqlConfig.PullSec = (int)numPull.Value;
            XqlConfig.DebounceMs = (int)numDebounce.Value;
            XqlConfig.HeartbeatSec = (int)numHeartbeat.Value;
            XqlConfig.LockTtlSec = (int)numLockTtl.Value;

            XqlConfig.Save(preferSidecar: true);
#if false
            if (chkSecure.Checked) 
                XqlGraphQLClient.Secrets.SaveApiKey(txtApiKey.Text); 
            else 
                XqlGraphQLClient.Secrets.Clear();
#endif
            if (apply)
            {
                btnSaveApply.Enabled = btnSave.Enabled = false;
                UseWaitCursor = true;
                
                BeginInvoke(() =>
                {
                    try
                    {
                        XqlAddIn.RestartRuntime();
                        MessageBox.Show(this, "Saved & Applied.", "XQLite", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(this, "Restart failed:\r\n" + ex, "XQLite", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    finally
                    {
                        UseWaitCursor = false;
                        btnSaveApply.Enabled = btnSave.Enabled = true;
                    }
                });
                return;
            }
            MessageBox.Show(this, "Saved.", "XQLite", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        // Validation
        private void RequiredValidator(object? sender, CancelEventArgs e)
        {
            if (sender is not TextBox tb) return;
            if (string.IsNullOrWhiteSpace(tb.Text))
            {
                e.Cancel = true;
                errors.SetError(tb, "필수 입력 항목입니다.");
            }
            else errors.SetError(tb, "");
        }
        private void OptionalNoOp(object? sender, CancelEventArgs e) { }
        private static int Clamp(int v, int lo, int hi) => Math.Max(lo, Math.Min(v, hi));

        // 디자이너 자동 생성 메서드는 비워둠
        private void InitializeComponent() { }
    }

    // placeholder helper
    internal static class TextBoxExt
    {
        public static void PlaceholderTextSafe(this TextBox tb, string text)
        {
#if NET6_0_OR_GREATER
            tb.PlaceholderText = text;
#else
            var gray = SystemColors.GrayText; var normal = tb.ForeColor;
            void SetPrompt() { if (string.IsNullOrEmpty(tb.Text) && !tb.Focused) { tb.Tag ??= text; tb.ForeColor = gray; tb.Text = text; } }
            void ClearPrompt() { if ((string?)tb.Tag == text && tb.Text == text) { tb.Text = string.Empty; tb.ForeColor = normal; } }
            tb.GotFocus += (_, __) => ClearPrompt();
            tb.LostFocus += (_, __) => SetPrompt();
            tb.TextChanged += (_, __) => { if (tb.ForeColor == gray && tb.Text != text) tb.ForeColor = normal; };
            tb.HandleCreated += (_, __) => SetPrompt();
#endif
        }
    }
    internal sealed class InlineButtonTextBox : UserControl
    {
        public TextBox TextBox { get; } = new TextBox();
        public Button TailButton { get; } = new Button();

        private int _buttonWidth;

        public InlineButtonTextBox(string buttonText = "Show", int buttonWidth = 52)
        {
            _buttonWidth = buttonWidth;

            // 바깥은 TextBox처럼 보이게
            BorderStyle = BorderStyle.FixedSingle;
            BackColor = SystemColors.Window;
            Margin = Padding = new Padding(0);
            Height = Math.Max(SystemFonts.MessageBoxFont.Height + 10, 24);
            MinimumSize = new Size(120, Height);

            // 버튼(오른쪽)
            TailButton.Text = buttonText;
            TailButton.AutoSize = false;
            TailButton.Width = _buttonWidth;
            TailButton.Dock = DockStyle.Right;
            TailButton.FlatStyle = FlatStyle.System;
            TailButton.TabStop = false;
            TailButton.Margin = new Padding(0);

            // 텍스트박스(한 줄, 테두리 제거)
            TextBox.BorderStyle = BorderStyle.None;
            TextBox.Multiline = false;            // 단일라인 유지
            TextBox.AutoSize = false;             // 높이 수동 제어
            TextBox.BackColor = SystemColors.Window;
            TextBox.TextAlign = HorizontalAlignment.Left;

            // 도킹 순서: Right → 수동 배치(TextBox는 Dock 안 씀)
            Controls.Add(TextBox);
            Controls.Add(TailButton);

            // 레이아웃 갱신 트리거
            SizeChanged += (_, __) => UpdateLayout();
            FontChanged += (_, __) => UpdateLayout();
            TextBox.FontChanged += (_, __) => UpdateLayout();
            HandleCreated += (_, __) => UpdateLayout();
        }

        private void UpdateLayout()
        {
            // 버튼은 Dock=Right로 이미 배치됨

            // 현재 폰트에서 한 줄 텍스트의 픽셀 높이 측정
            // 약간의 여유(+2)로 클리핑 방지
            int textPx = TextRenderer.MeasureText("Ag", TextBox.Font).Height;
            int tbH = Math.Max(textPx + 2, 16);

            // 컨테이너 중앙에 오도록 Top 계산
            int top = Math.Max((ClientSize.Height - tbH) / 2, 1);

            // 좌우 여백
            int leftPad = 6;
            int rightPad = 6;

            // 버튼을 뺀 가용 폭
            int availW = ClientSize.Width - TailButton.Width - leftPad - rightPad - 1;
            if (availW < 30) availW = 30;

            TextBox.SetBounds(leftPad, top, availW, tbH);
            // 배경/테두리 색 보정(모던 룩)
            BackColor = Enabled ? SystemColors.Window : SystemColors.Control;
            TextBox.BackColor = BackColor;
        }
    }

}
