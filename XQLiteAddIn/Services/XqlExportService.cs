using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

#if false
namespace XQLite.AddIn
{
    internal static class XqlExportService
    {
        internal static async Task ExportSnapshotAsync(long since = 0, string? targetDir = null, bool csv = false)
        {
            const string q = "query($since:Long){ rows(since_version:$since){ table, rows, max_row_version } }";
            var resp = await XqlGraphQLClient.QueryAsync<XqlUpsert.RowsResp>(q, new { since });
            var blocks = resp.Data?.rows; if (blocks is null || blocks.Length == 0) return;

            string dir = targetDir ?? Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "XQLiteSnapshots");
            Directory.CreateDirectory(dir);

            foreach (var blk in blocks)
            {
                var table = blk.table;
                var ts = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                if (!csv)
                {
                    // JSON
                    var path = Path.Combine(dir, $"{table}_{ts}.json");
                    var json = XqlJson.Serialize(blk.rows ?? Array.Empty<Dictionary<string, object?>>(), true);

                    File.WriteAllText(path, json, Encoding.UTF8);
                }
                else
                {
                    // CSV: 컬럼 헤더 유니온
                    var rows = blk.rows ?? Array.Empty<Dictionary<string, object?>>();
                    var headers = rows.SelectMany(r => r.Keys).Distinct(StringComparer.OrdinalIgnoreCase).ToArray();
                    var path = Path.Combine(dir, $"{table}_{ts}.csv");
                    using var sw = new StreamWriter(path, false, Encoding.UTF8);
                    await sw.WriteLineAsync(string.Join(",", headers.Select(EscapeCsv)));
                    foreach (var r in rows)
                    {
                        var line = string.Join(",", headers.Select(h => EscapeCsv(r.TryGetValue(h, out var v) ? v : null)));
                        await sw.WriteLineAsync(line);
                    }
                }
            }
        }

        private static string EscapeCsv(object? v)
        {
            if (v is null) return string.Empty;
            var s = Convert.ToString(v)?.Replace("\r", " ").Replace("\n", " ") ?? string.Empty;
            if (s.Contains(',') || s.Contains('"')) s = '"' + s.Replace("\"", "\"\"") + '"';
            return s;
        }
    }
}
#endif