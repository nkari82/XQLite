using ExcelDna.Integration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace XQLite.AddIn
{
    internal static class XqlRecoverService
    {
        internal static async Task RecoverAllTablesAsync(int batch = 500)
        {
            var app = (Excel.Application)ExcelDnaUtil.Application;
            foreach (Excel.Worksheet ws in app.Worksheets)
            {
                if (ws.ListObjects.Count == 0) continue;
                foreach (Excel.ListObject lo in ws.ListObjects)
                {
                    var body = lo.DataBodyRange; if (body == null) continue;
                    var header = lo.HeaderRowRange; int colCount = header.Columns.Count;
                    var headerArr = (object[,])header.Value2;
                    string[] cols = Enumerable.Range(1, colCount).Select(c => Convert.ToString(headerArr[1, c]) ?? $"C{c}").ToArray();
                    string tableName = string.IsNullOrWhiteSpace(lo.Name) ? ws.Name : lo.Name;

                    int total = body.Rows.Count; int idx = 1;
                    while (idx <= total)
                    {
                        int take = Math.Min(batch, total - idx + 1);
                        var seg = body.Range[body.Cells[idx, 1], body.Cells[idx + take - 1, colCount]];
                        var arr = (object[,])seg.Value2;
                        var rows = new List<Dictionary<string, object?>>(take);

                        for (int r = 1; r <= take; r++)
                        {
                            var d = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
                            for (int c = 1; c <= colCount; c++)
                            {
                                var v = arr[r, c]; if (v is string s && string.IsNullOrWhiteSpace(s)) v = null;
                                d[cols[c - 1]] = v;
                            }
                            rows.Add(d);
                        }

                        const string m = @"mutation ($table:String!,$rows:[JSON!]!){
  upsertRows(table:$table, rows:$rows){ affected, errors{code,message}, max_row_version }
}";
                        try
                        {
                            var resp = await XqlGraphQLClient.MutateAsync<XqlUpsert.UpsertResp>(m, new { table = tableName, rows });
                            var data = resp.Data?.upsertRows;
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"recover {tableName} failed: {ex.Message}");
                        }

                        idx += take;
                    }
                }
            }
        }
    }
}