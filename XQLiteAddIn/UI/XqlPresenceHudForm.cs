using System;
using System.Drawing;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace XQLite.AddIn
{
    public sealed class XqlPresenceHudForm : Form
    {
        private static XqlPresenceHudForm? _inst; 
        
        internal static void ShowSingleton() 
        { 
            if (_inst == null || _inst.IsDisposed) 
                _inst = new XqlPresenceHudForm(); 
            _inst.Show(); 
            _inst.BringToFront(); 
        }

        private ListView lv = new(); private Timer auto = new() { Interval = 2000 };

        public XqlPresenceHudForm()
        {
            Text = "XQLite Presence"; StartPosition = FormStartPosition.Manual; Left = 20; Top = 20; Width = 520; Height = 320;
            lv.View = View.Details; lv.FullRowSelect = true; lv.Dock = DockStyle.Fill;
            lv.Columns.AddRange(new[] { new ColumnHeader { Text = "Nickname", Width = 140 }, new ColumnHeader { Text = "Sheet/Cell", Width = 200 }, new ColumnHeader { Text = "When", Width = 140 } });
            Controls.Add(lv); auto.Tick += async (_, __) => await RefreshAsync(); auto.Start(); Load += async (_, __) => await RefreshAsync();
        }

        private async Task RefreshAsync()
        {
            try
            {
                const string q = "query{ presence{ nickname sheet cell updated_at } }";
                var resp = await XqlGraphQLClient.QueryAsync<PresenceResp>(q, null);
                var list = resp.Data?.presence ?? Array.Empty<PresenceItem>();
                lv.BeginUpdate(); lv.Items.Clear();
                foreach (var p in list)
                    lv.Items.Add(new ListViewItem(new[] { p.nickname ?? "", string.Format("{0}/{1}", p.sheet, p.cell), p.updated_at ?? "" }));
                lv.EndUpdate();
            }
            catch { /* 서버가 지원 안하면 조용히 */ }
        }

        private sealed class PresenceResp { public PresenceItem[]? presence { get; set; } }
        private sealed class PresenceItem { public string? nickname { get; set; } public string? sheet { get; set; } public string? cell { get; set; } public string? updated_at { get; set; } }
    }
}