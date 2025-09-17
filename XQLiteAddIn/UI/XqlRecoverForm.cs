using ExcelDna.Integration;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace XQLite.AddIn
{
    public sealed class XqlRecoverForm : Form
    {
        private static XqlRecoverForm? _inst;
        public static void ShowSingleton() { if (_inst == null || _inst.IsDisposed) _inst = new XqlRecoverForm(); _inst.Show(); _inst.BringToFront(); }

        private ProgressBar pb = new() { Dock = DockStyle.Top, Height = 18, Minimum = 0, Maximum = 100 };
        private Label lbl = new() { Dock = DockStyle.Top, Height = 22, Text = "Ready" };
        private Button btnRun = new() { Text = "Recover (Start)", Dock = DockStyle.Left, Width = 140 };
        private Button btnCancel = new() { Text = "Cancel", Dock = DockStyle.Right, Width = 100, Enabled = false };
        private NumericUpDown numBatch = new() { Minimum = 100, Maximum = 5000, Value = 1000, Dock = DockStyle.Top, Width = 120 };
        private NumericUpDown numDegree = new() { Minimum = 1, Maximum = 8, Value = 2, Dock = DockStyle.Top, Width = 120 };
        private CancellationTokenSource? _cts;

        public XqlRecoverForm()
        {
            Text = "XQLite Recover"; StartPosition = FormStartPosition.CenterScreen; Width = 520; Height = 160;
            var panel = new FlowLayoutPanel { Dock = DockStyle.Top, Height = 32, FlowDirection = FlowDirection.LeftToRight };
            panel.Controls.Add(new Label { Text = "Batch", AutoSize = true, Padding = new Padding(0, 6, 8, 0) }); panel.Controls.Add(numBatch);
            panel.Controls.Add(new Label { Text = "Parallel", AutoSize = true, Padding = new Padding(12, 6, 8, 0) }); panel.Controls.Add(numDegree);
            var bottom = new Panel { Dock = DockStyle.Bottom, Height = 36 }; bottom.Controls.Add(btnRun); bottom.Controls.Add(btnCancel);
            Controls.Add(bottom); Controls.Add(panel); Controls.Add(lbl); Controls.Add(pb);
            btnRun.Click += async (_, __) => await RunAsync();
            btnCancel.Click += (_, __) => _cts?.Cancel();
        }

        private async Task RunAsync()
        {
            if (_cts != null) return;
            _cts = new CancellationTokenSource(); btnCancel.Enabled = true; btnRun.Enabled = false; pb.Value = 0; lbl.Text = "Collecting tables...";
            try
            {
                var tables = XqlFastTableReader.Collect();
                int totalRows = tables.Sum(t => t.BodyRows);
                int doneRows = 0; int failures = 0;

                await Task.WhenAll(tables.Select(t => Task.Run(async () =>
                {
                    string tableName = XqlTableNameMap.Map(t.ListObjectName, t.WorksheetName);
                    var headers = XqlFastTableReader.ReadHeader(t);
                    int adaptive = XqlAdaptiveBatcher.PickBatchSize(t.BodyRows, (int)numBatch.Minimum, (int)numBatch.Maximum);
                    int batch = Math.Min((int)numBatch.Value, adaptive);

                    int idx = 1; var (scope, rate) = XqlPerf.Scope("Recover:" + tableName, tableName);
                    using (scope)
                    {
                        while (idx <= t.BodyRows)
                        {
                            _cts!.Token.ThrowIfCancellationRequested();
                            int take = Math.Min(batch, t.BodyRows - (idx - 1));
                            var rows = XqlFastTableReader.ReadBodyChunks(t, headers, idx, take).ToList();
                            // 대강의 바이트 추정(키/값 문자열 길이 합)
                            long approxBytes = rows.Sum(r => r.Sum(kv => (kv.Key.Length + (kv.Value?.ToString()?.Length ?? 0))));

                            const string m = @"mutation ($table:String!,$rows:[JSON!]!){ upsertRows(table:$table, rows:$rows){ affected, errors{code,message}, max_row_version } }";
                            try
                            {
                                var resp = await XqlGraphQLClient.MutateAsync<XqlUpsert.UpsertResp>(m, new { table = tableName, rows });
                                var data = resp.Data?.upsertRows;
                                if (data?.errors?.Length > 0) Interlocked.Add(ref failures, data.errors.Length);
                            }
                            catch { Interlocked.Increment(ref failures); }

                            Interlocked.Add(ref doneRows, take);
                            idx += take;
                            rate(approxBytes);
                            UpdateUi(doneRows, totalRows, $"{tableName}: {doneRows}/{totalRows} rows");
                        }
                    }
                }, _cts.Token)).ToArray());

                UpdateUi(totalRows, totalRows, failures == 0 ? "Recover done" : "Recover done with errors: " + failures);
            }
            catch (OperationCanceledException) { UpdateUi(pb.Maximum, pb.Maximum, "Cancelled"); }
            catch (Exception ex) { UpdateUi(pb.Value, pb.Maximum, "Failed: " + ex.Message); }
            finally { _cts?.Dispose(); _cts = null; btnCancel.Enabled = false; btnRun.Enabled = true; }
        }

        private void UpdateUi(int cur, int total, string text) { if (total <= 0) total = 1; int pct = Math.Min(100, Math.Max(0, (int)(100.0 * cur / total))); pb.Value = pct; lbl.Text = text; }

        private static List<TableSlice> CollectTables()
        {
            var app = (Excel.Application)ExcelDnaUtil.Application; var list = new List<TableSlice>();
            foreach (Excel.Worksheet ws in app.Worksheets)
            {
                if (ws.ListObjects.Count == 0) continue;
                foreach (Excel.ListObject lo in ws.ListObjects)
                {
                    var body = lo.DataBodyRange; if (body == null) continue; var header = lo.HeaderRowRange; int cols = header.Columns.Count;
                    list.Add(new TableSlice(ws.Name, lo.Name, body.Row, body.Rows.Count, cols, body.Column));
                }
            }
            return list;
        }

        private static List<Dictionary<string, object?>> ReadRows(TableSlice t, int start1, int take)
        {
            var app = (Excel.Application)ExcelDnaUtil.Application; var ws = (Excel.Worksheet)app.Worksheets[t.WorksheetName];
            var lo = ws.ListObjects[t.ListObjectName]; var header = lo.HeaderRowRange; var body = lo.DataBodyRange!;
            int colCount = header.Columns.Count; var headerArr = (object[,])header.Value2; var cols = Enumerable.Range(1, colCount).Select(c => Convert.ToString(headerArr[1, c]) ?? "C" + c).ToArray();
            var seg = body.Range[body.Cells[start1, 1], body.Cells[start1 + take - 1, colCount]]; var arr = (object[,])seg.Value2;
            var rows = new List<Dictionary<string, object?>>(take);
            for (int r = 1; r <= take; r++)
            {
                var d = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
                for (int c = 1; c <= colCount; c++) { var v = arr[r, c]; if (v is string s && string.IsNullOrWhiteSpace(s)) v = null; d[cols[c - 1]] = v; }
                rows.Add(d);
            }
            return rows;
        }

        private sealed record TableSlice(string WorksheetName, string ListObjectName, int BodyTop, int BodyRows, int Cols, int BodyLeft);
    }
}