using ExcelDna.Integration;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace XQLite.AddIn
{
    public sealed class XqlRecoverForm : Form
    {
        public sealed record Slice(string WorksheetName, string ListObjectName, int BodyTop, int BodyRows, int Cols, int BodyLeft);

        private static XqlRecoverForm? _inst;
        internal static void ShowSingleton()
        {
            if (_inst == null || _inst.IsDisposed)
                _inst = new XqlRecoverForm();
            _inst.Show();
            _inst.BringToFront();
        }

        private readonly ProgressBar pb = new() { Dock = DockStyle.Top, Height = 18, Minimum = 0, Maximum = 100 };
        private readonly Label lbl = new() { Dock = DockStyle.Top, Height = 22, Text = "Ready" };
        private readonly Button btnRun = new() { Text = "Recover (Start)", Dock = DockStyle.Left, Width = 140 };
        private readonly Button btnCancel = new() { Text = "Cancel", Dock = DockStyle.Right, Width = 100, Enabled = false };
        private readonly NumericUpDown numBatch = new() { Minimum = 100, Maximum = 5000, Value = 1000, Dock = DockStyle.Top, Width = 120 };
        private readonly NumericUpDown numDegree = new() { Minimum = 1, Maximum = 8, Value = 2, Dock = DockStyle.Top, Width = 120 };
        private CancellationTokenSource? _cts;

        public XqlRecoverForm()
        {
            Text = "XQLite Recover";
            StartPosition = FormStartPosition.CenterScreen;
            Width = 520; Height = 160;

            var panel = new FlowLayoutPanel { Dock = DockStyle.Top, Height = 32, FlowDirection = FlowDirection.LeftToRight };
            panel.Controls.Add(new Label { Text = "Batch", AutoSize = true, Padding = new Padding(0, 6, 8, 0) }); panel.Controls.Add(numBatch);
            panel.Controls.Add(new Label { Text = "Parallel", AutoSize = true, Padding = new Padding(12, 6, 8, 0) }); panel.Controls.Add(numDegree);

            var bottom = new Panel { Dock = DockStyle.Bottom, Height = 36 };
            bottom.Controls.Add(btnRun); bottom.Controls.Add(btnCancel);

            Controls.Add(bottom); Controls.Add(panel); Controls.Add(lbl); Controls.Add(pb);

            btnRun.Click += async (_, __) => await RunAsync();
            btnCancel.Click += (_, __) => _cts?.Cancel();
        }

        internal static int PickBatchSize(int rowCount, int min = 200, int max = 2000)
        {
            if (rowCount > 20000) return min;
            if (rowCount > 5000) return Math.Min(1000, max);
            return Math.Min(max, 1500);
        }

        private async Task RunAsync()
        {
            if (_cts != null) return;

            // Backend null 가드
            if (XqlAddIn.Backend is not IXqlBackend be)
                return;


            _cts = new CancellationTokenSource();
            btnCancel.Enabled = true; btnRun.Enabled = false; pb.Value = 0; lbl.Text = "Collecting tables...";
            try
            {

                var tables = Collect();
                int totalRows = tables.Sum(t => t.BodyRows);
                int doneRows = 0; int failures = 0;

                // 테이블 단위 병렬(최대 numDegree)
                using var sem = new SemaphoreSlim((int)numDegree.Value);
                var tasks = tables.Select(async t =>
                {
                    await sem.WaitAsync(_cts.Token).ConfigureAwait(false);
                    try
                    {
                        string tableName = XqlTableNameMap.Map(t.ListObjectName, t.WorksheetName);
                        var headers = ReadHeader(t);

                        int adaptive = PickBatchSize(t.BodyRows, (int)numBatch.Minimum, (int)numBatch.Maximum);
                        int batch = Math.Min((int)numBatch.Value, adaptive);

                        int idx = 1;
                        var (scope, rate) = Scope("Recover:" + tableName, tableName);
                        using (scope)
                        {
                            while (idx <= t.BodyRows)
                            {
                                _cts!.Token.ThrowIfCancellationRequested();
                                int take = Math.Min(batch, t.BodyRows - (idx - 1));

                                var rows = ReadBodyChunks(t, headers, idx, take).ToList();

                                // 대강의 바이트 추정(키/값 문자열 길이 합)
                                long approxBytes = rows.Sum(r => r.Sum(kv => (kv.Key.Length + (kv.Value?.ToString()?.Length ?? 0))));

                                try
                                {
                                    var ok = await be.UpsertRows(tableName, rows, _cts.Token).ConfigureAwait(false);
                                    if (!ok) Interlocked.Increment(ref failures);
                                }
                                catch
                                {
                                    Interlocked.Increment(ref failures);
                                }

                                Interlocked.Add(ref doneRows, take);
                                idx += take;
                                rate(approxBytes);
                                UpdateUi(doneRows, totalRows, $"{tableName}: {doneRows}/{totalRows} rows");
                            }
                        }
                    }
                    finally
                    {
                        sem.Release();
                    }
                }).ToArray();

                await Task.WhenAll(tasks).ConfigureAwait(false);

                UpdateUi(totalRows, totalRows, failures == 0 ? "Recover done" : "Recover done with errors: " + failures);
            }
            catch (OperationCanceledException) { UpdateUi(pb.Maximum, pb.Maximum, "Cancelled"); }
            catch (Exception ex) { UpdateUi(pb.Value, pb.Maximum, "Failed: " + ex.Message); }
            finally { _cts?.Dispose(); _cts = null; btnCancel.Enabled = false; btnRun.Enabled = true; }
        }

        private void UpdateUi(int cur, int total, string text)
        {
            if (total <= 0) total = 1;
            int pct = Math.Min(100, Math.Max(0, (int)(100.0 * cur / total)));
            pb.Value = pct; lbl.Text = text;
        }

        // ─────────────────────────────────────────────────────────
        // Excel helpers
        // ─────────────────────────────────────────────────────────
        internal static List<Slice> Collect()
        {
            var app = (Excel.Application)ExcelDnaUtil.Application;
            var list = new List<Slice>();
            foreach (Excel.Worksheet ws in app.Worksheets)
            {
                if (ws.ListObjects.Count == 0) continue;

                foreach (Excel.ListObject lo in ws.ListObjects)
                {
                    var body = lo.DataBodyRange;
                    if (body == null) continue;

                    var header = lo.HeaderRowRange;
                    int cols = header.Columns.Count;
                    list.Add(new Slice(ws.Name, lo.Name, body.Row, body.Rows.Count, cols, body.Column));
                }
            }
            return list;
        }

        internal static string[] ReadHeader(Slice s)
        {
            var app = (Excel.Application)ExcelDnaUtil.Application;
            var ws = (Excel.Worksheet)app.Worksheets[s.WorksheetName];
            var lo = ws.ListObjects[s.ListObjectName];
            var header = lo.HeaderRowRange;
            int colCount = header.Columns.Count;
            var arr = (object[,])header.Value2;
            var headers = new string[colCount];
            for (int c = 1; c <= colCount; c++)
                headers[c - 1] = Convert.ToString(arr[1, c]) ?? $"C{c}";
            return headers;
        }

        internal static IEnumerable<Dictionary<string, object?>> ReadBodyChunks(Slice s, string[] headers, int startRow1, int take)
        {
            var app = (Excel.Application)ExcelDnaUtil.Application;
            var ws = (Excel.Worksheet)app.Worksheets[s.WorksheetName];
            var lo = ws.ListObjects[s.ListObjectName];
            var body = lo.DataBodyRange!;
            int colCount = headers.Length;
            int endRow1 = Math.Min(startRow1 + take - 1, s.BodyRows);
            var seg = body.Range[body.Cells[startRow1, 1], body.Cells[endRow1, colCount]];
            var arr = (object[,])seg.Value2;
            for (int r = 1; r <= endRow1 - startRow1 + 1; r++)
            {
                var d = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
                for (int c = 1; c <= colCount; c++)
                {
                    var v = arr[r, c];
                    if (v is string ss)
                    {
                        ss = ss.Trim();
                        if (string.IsNullOrEmpty(ss))
                            v = null;
                        else
                            v = ss;
                    }
                    d[headers[c - 1]] = v; // 숫자/날짜 등은 Value2 원본 유지
                }
                yield return d;
            }
        }

        internal static (IDisposable scope, Action<long> done) Scope(string name, string table = "*")
        {
            var sw = Stopwatch.StartNew();
            return (new ScopeDisposable(() =>
            {
                var ms = sw.ElapsedMilliseconds; XqlLog.Info($"{name} took {ms} ms");
            }), bytes =>
            {
                var ms = Math.Max(1, sw.ElapsedMilliseconds);
                var kbps = (bytes / 1024.0) / (ms / 1000.0);
                XqlLog.Info($"{name}: ~{kbps:F1} KB/s (approx) on {table}");
            }
            );
        }

        private sealed class ScopeDisposable : IDisposable
        {
            private readonly Action _on;
            public ScopeDisposable(Action on) { _on = on; }
            public void Dispose() { _on(); }
        }
    }
}
