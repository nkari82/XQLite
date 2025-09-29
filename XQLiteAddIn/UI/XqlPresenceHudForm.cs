// XqlPresenceHudForm.cs
using System;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Timer = System.Windows.Forms.Timer;

namespace XQLite.AddIn
{
    public sealed class XqlPresenceHudForm : Form
    {
        private static XqlPresenceHudForm? _inst;
        internal static void ShowSingleton()
        {
            if (_inst == null || _inst.IsDisposed) _inst = new XqlPresenceHudForm();
            _inst.Show();
            _inst.BringToFront();
        }

        private readonly ListView _lv = new() { View = View.Details, Dock = DockStyle.Fill, FullRowSelect = true };
        private readonly Timer _auto = new() { Interval = 3000 }; // 3s 주기
        private volatile int _refreshing; // 0:idle, 1:busy
        private CancellationTokenSource? _cts; // 폼 종료 시 취소용

        public XqlPresenceHudForm()
        {
            Text = "XQLite Presence";
            StartPosition = FormStartPosition.Manual;
            Left = 20; Top = 20; Width = 560; Height = 360;

            _lv.Columns.AddRange(
            [
                new ColumnHeader { Text = "Nickname", Width = 140 },
                new ColumnHeader { Text = "Sheet/Cell", Width = 220 },
                new ColumnHeader { Text = "When (UTC)", Width = 160 }
            ]);

            Controls.Add(_lv);

            Load += async (_, __) => await RefreshList().ConfigureAwait(false);
            FormClosed += (_, __) => { try { _cts?.Cancel(); } catch { } };
            _auto.Tick += async (_, __) => await RefreshList().ConfigureAwait(false);
            _auto.Start();
        }

        private async Task RefreshList()
        {
            if (XqlAddIn.Backend is not IXqlBackend be) return;

            if (Interlocked.Exchange(ref _refreshing, 1) == 1) return;

            CancellationTokenSource? prev = _cts;
            _cts = new CancellationTokenSource(TimeSpan.FromSeconds(5));
            try
            {
                prev?.Dispose(); // 이전 CTS 정리

                var list = await be.FetchPresence(_cts.Token).ConfigureAwait(false);

                // UI 갱신
                if (IsHandleCreated && !IsDisposed && list != null)
                {
                    BeginInvoke(new Action(() =>
                    {
                        if (IsDisposed) return;
                        _lv.BeginUpdate();
                        try
                        {
                            _lv.Items.Clear();
                            foreach (var p in list)
                            {
                                var when = p.updated_at ?? "";
                                var where = string.IsNullOrWhiteSpace(p.sheet) && string.IsNullOrWhiteSpace(p.cell)
                                            ? ""
                                            : $"{p.sheet}/{p.cell}";
                                _lv.Items.Add(new ListViewItem(new[] { p.nickname ?? "", where, when }));
                            }
                        }
                        finally { _lv.EndUpdate(); }
                    }));
                }
            }
            catch { /* 네트워크 일시 오류 무음 */ }
            finally
            {
                Interlocked.Exchange(ref _refreshing, 0);
            }
        }

    }
}
