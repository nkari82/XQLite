using System;
using System.Drawing;
using System.Windows.Forms;

namespace XQLite.AddIn
{
    public sealed class XqlCommandPaletteForm : Form
    {
        private static XqlCommandPaletteForm? _inst;
        public static void ShowSingleton() 
        { 
            if (_inst == null || _inst.IsDisposed) 
                _inst = new XqlCommandPaletteForm(); 
            _inst.Show(); 
            _inst.BringToFront(); 
        }

        private TextBox txt = new() { Dock = DockStyle.Top, Font = new Font("Segoe UI", 11f) };
        private ListView lv = new() { Dock = DockStyle.Fill, View = View.Details, FullRowSelect = true, HideSelection = false };

        private readonly (string title, string hint, Action run)[] _items =
        [
            ("Commit", "Ctrl+Shift+C", XqlCommands.CommitCommand),
            ("Recover", "Ctrl+Shift+R", XqlCommands.RecoverCommand),
            ("Inspector", "Ctrl+Shift+I", XqlCommands.InspectorCommand),
            ("Export Snapshot", "Ctrl+Shift+E", XqlCommands.ExportSnapshotCommand),
            ("Presence HUD", "Ctrl+Shift+P", XqlCommands.PresenceCommand),
            ("Schema Explorer", "Ctrl+Shift+S", XqlCommands.SchemaCommand),
            ("Export Diagnostics", "Ctrl+Shift+D", ()=> XqlCommands.ExportDiagnosticsCommand()),
            ("Open Config", "(Ribbon)", XqlCommands.ConfigCommand)
        ];

        public XqlCommandPaletteForm()
        {
            StartPosition = FormStartPosition.CenterScreen; Width = 520; Height = 360; Text = "XQLite Command";
            lv.Columns.AddRange([new ColumnHeader { Text = "Command", Width = 260 }, new ColumnHeader { Text = "Shortcut", Width = 180 }]);
            Controls.Add(lv); Controls.Add(txt);
            txt.TextChanged += (_, __) => RefreshList();
            lv.DoubleClick += (_, __) => RunSelected();
            KeyPreview = true;
            KeyDown += (s, e) => { if (e.KeyCode == Keys.Enter) { RunSelected(); e.Handled = true; } else if (e.KeyCode == Keys.Escape) { Close(); } };
            Load += (_, __) => { RefreshList(); txt.Focus(); txt.SelectAll(); };
        }

        private void RefreshList()
        {
            string q = txt.Text.Trim().ToLowerInvariant();
            lv.BeginUpdate(); lv.Items.Clear();
            foreach (var it in _items)
            {
                if (!string.IsNullOrEmpty(q) && !it.title.ToLowerInvariant().Contains(q)) continue;
                lv.Items.Add(new ListViewItem([it.title, it.hint]) { Tag = it.run });
            }
            lv.EndUpdate(); if (lv.Items.Count > 0) lv.Items[0].Selected = true;
        }

        private void RunSelected()
        {
            if (lv.SelectedItems.Count == 0) return;
            var act = lv.SelectedItems[0].Tag as Action; act?.Invoke(); Close();
        }
    }
}