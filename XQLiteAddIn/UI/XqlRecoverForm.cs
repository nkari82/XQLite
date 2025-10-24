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

            btnRun.Click += async (_, __) => await Run();
            btnCancel.Click += (_, __) => _cts?.Cancel();
        }

        internal static int PickBatchSize(int rowCount, int min = 200, int max = 2000)
        {
            if (rowCount > 20000) return min;
            if (rowCount > 5000) return Math.Min(1000, max);
            return Math.Min(max, 1500);
        }

        private async Task Run()
        {
            if (_cts != null) return;

            if (XqlAddIn.Backend is not IXqlBackend be) return;

            _cts = new CancellationTokenSource();
            btnCancel.Enabled = true; btnRun.Enabled = false; pb.Value = 0; lbl.Text = "Collecting tables...";
            try
            {
                var tables = Collect();
                int totalRows = tables.Sum(t => t.BodyRows);
                int doneRows = 0; int failures = 0;

                using var sem = new SemaphoreSlim((int)numDegree.Value);
                var tasks = tables.Select(async t =>
                {
                    await sem.WaitAsync(_cts.Token).ConfigureAwait(false);
                    try
                    {
                        string tableName = XqlTableNameMap.Map(t.ListObjectName, t.WorksheetName);
                        var headers = ReadHeader(t);

                        // 키 컬럼 인덱스(id 기본)
                        int keyColIdx1 = Math.Max(1, Array.FindIndex(headers, h => string.Equals(h, "id", StringComparison.OrdinalIgnoreCase)) + 1);

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

                                var chunk = ReadBodyChunkWithRowIndex(t, headers, idx, take).ToList();

                                // 두 갈래로 분류
                                var rowsForUpsert = new List<Dictionary<string, object?>>(chunk.Count);
                                var cellsForInsert = new List<EditCell>(chunk.Count * Math.Max(1, headers.Length - 1));
                                var tempKeyToExcelRow = new Dictionary<string, int>(StringComparer.Ordinal);

                                foreach (var item in chunk)
                                {
                                    var row = item.Data;
                                    // id 유무 판단
                                    object? idVal;
                                    row.TryGetValue(headers[keyColIdx1 - 1], out idVal);
                                    var idStr = XqlCommon.Canonicalize(idVal) ?? "";

                                    // 완전 빈 행 스킵 (id도 비고 나머지 비어있음)
                                    bool anyData = row.Any(kv => !string.Equals(kv.Key, headers[keyColIdx1 - 1], StringComparison.OrdinalIgnoreCase)
                                                              && kv.Value != null
                                                              && !string.IsNullOrWhiteSpace(Convert.ToString(kv.Value)));
                                    if (!anyData && string.IsNullOrWhiteSpace(idStr)) continue;

                                    if (!string.IsNullOrWhiteSpace(idStr))
                                    {
                                        // 기존행 → upsertRows
                                        rowsForUpsert.Add(row);
                                    }
                                    else
                                    {
                                        // 신규행 → upsertCells (임시키 = "-<엑셀실제행번호>")
                                        string tempKey = "-" + (t.BodyTop + item.RowOffset0).ToString();
                                        tempKeyToExcelRow[tempKey] = t.BodyTop + item.RowOffset0;
                                        foreach (var (k, v) in row)
                                        {
                                            if (string.Equals(k, headers[keyColIdx1 - 1], StringComparison.OrdinalIgnoreCase)) continue;
                                            cellsForInsert.Add(new EditCell(tableName, tempKey, k, v));
                                        }
                                    }
                                }

                                // 대강의 바이트 추정(키/값 문자열 길이 합)
                                long approxBytes =
                                    rowsForUpsert.Sum(r => r.Sum(kv => (kv.Key.Length + (kv.Value?.ToString()?.Length ?? 0))))
                                    + cellsForInsert.Count * 12;

                                try
                                {
                                    // 1) 기존행 반영
                                    if (rowsForUpsert.Count > 0)
                                    {
                                        var res = await be.UpsertRows(tableName, rowsForUpsert, _cts.Token).ConfigureAwait(false);
                                        if (res?.Errors is { Count: > 0 })
                                        {
                                            Interlocked.Add(ref failures, res.Errors.Count);
                                            XqlLog.Warn($"Recover[{tableName}] upsertRows errors: " + string.Join("; ", res.Errors));
                                        }
                                    }

                                    // 2) 신규행 반영 + 배정된 id를 시트에 기록
                                    if (cellsForInsert.Count > 0)
                                    {
                                        var res2 = await be.UpsertCells(cellsForInsert, _cts.Token).ConfigureAwait(false);
                                        if (res2?.Errors is { Count: > 0 })
                                        {
                                            Interlocked.Add(ref failures, res2.Errors.Count);
                                            XqlLog.Warn($"Recover[{tableName}] upsertCells errors: " + string.Join("; ", res2.Errors));
                                        }

                                        // assigned 반영 (신형/구형 모두 지원)
                                        if (res2?.Assigned != null && res2.Assigned.Count > 0)
                                        {
                                            var app = (Excel.Application)ExcelDnaUtil.Application;
                                            Excel.Worksheet? ws = null;

                                            try
                                            {
                                                try { ws = app.Worksheets[t.WorksheetName] as Excel.Worksheet; }
                                                catch (Exception ex) { XqlLog.Warn($"Recover: worksheet access failed: {ex.Message}"); ws = null; }

                                                if (ws != null)
                                                {
                                                    foreach (var a in res2.Assigned)
                                                    {
                                                        if (a == null) continue;

                                                        // 신형: table/temp_row_key/new_id
                                                        var tempKey = GetProp(a, "temp_row_key");
                                                        var newId = GetProp(a, "new_id");

                                                        // 구형 폴백: client_row/row_key (client_row는 없음이므로 생략)
                                                        if (string.IsNullOrWhiteSpace(tempKey))
                                                            tempKey = GetProp(a, "client_temp") ?? GetProp(a, "client_row") ?? "";
                                                        if (string.IsNullOrWhiteSpace(newId))
                                                            newId = GetProp(a, "row_key");

                                                        if (string.IsNullOrWhiteSpace(tempKey) || string.IsNullOrWhiteSpace(newId)) continue;

                                                        if (tempKeyToExcelRow.TryGetValue(tempKey!, out var excelRow))
                                                        {
                                                            Excel.Range? keyCell = null;
                                                            try
                                                            {
                                                                keyCell = (Excel.Range)ws.Cells[excelRow, t.BodyLeft + keyColIdx1 - 1];
                                                                keyCell.Value2 = newId;
                                                            }
                                                            catch (Exception ex)
                                                            {
                                                                XqlLog.Warn($"Recover assigned id write failed for tempKey={tempKey}: {ex.Message}");
                                                            }
                                                            finally { XqlCommon.ReleaseCom(keyCell); }
                                                        }
                                                    }
                                                }
                                            }
                                            finally
                                            {
                                                XqlCommon.ReleaseCom(ws);
                                                // ExcelDnaUtil.Application은 전역 어플리케이션 객체이므로 보통 Release하지 않습니다.
                                            }
                                        }
                                    }
                                }
                                catch (Exception ex)
                                {
                                    Interlocked.Increment(ref failures);
                                    XqlLog.Warn($"Recover[{tableName}] chunk failed: {ex.Message}");
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

        private static string? GetProp(object obj, string name)
        {
            try
            {
                var t = obj.GetType();
                var p = t.GetProperty(name);
                if (p != null) return Convert.ToString(p.GetValue(obj));
            }
            catch { }
            return null;
        }

        private void UpdateUi(int cur, int total, string text)
        {
            if (!IsHandleCreated || IsDisposed) return;
            BeginInvoke(new Action(() =>
            {
                if (IsDisposed) return;
                if (total <= 0) total = 1;
                int pct = Math.Min(100, Math.Max(0, (int)(100.0 * cur / total)));
                pb.Value = pct; lbl.Text = text;
            }));
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
            {
                var raw = Convert.ToString(arr[1, c]) ?? $"C{c}";
                raw = raw.Trim();
                headers[c - 1] = string.IsNullOrWhiteSpace(raw) ? $"C{c}" : raw;
            }
            return headers;
        }

        internal sealed class RowRead
        {
            public Dictionary<string, object?> Data { get; set; } = default!;
            public int RowOffset0 { get; set; } // body 내 0-based offset
        }

        internal static IEnumerable<RowRead> ReadBodyChunkWithRowIndex(Slice s, string[] headers, int startRow1, int take)
        {
            var app = (Excel.Application)ExcelDnaUtil.Application;
            var ws = (Excel.Worksheet)app.Worksheets[s.WorksheetName];
            var lo = ws.ListObjects[s.ListObjectName];
            var body = lo.DataBodyRange!;
            int colCount = headers.Length;
            int endRow1 = Math.Min(startRow1 + take - 1, s.BodyRows);
            var seg = body.Range[body.Cells[startRow1, 1], body.Cells[endRow1, colCount]];
            var arr = (object[,])seg.Value2;
            int rows = endRow1 - startRow1 + 1;

            for (int r = 1; r <= rows; r++)
            {
                var d = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
                for (int c = 1; c <= colCount; c++)
                {
                    var v = arr[r, c];
                    if (v is string ss)
                    {
                        ss = ss.Trim();
                        v = string.IsNullOrEmpty(ss) ? null : ss;
                    }
                    d[headers[c - 1]] = v; // 숫자/날짜 등은 Value2 원본 유지
                }
                yield return new RowRead { Data = d, RowOffset0 = (startRow1 - 1) + (r - 1) };
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
