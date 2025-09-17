using ExcelDna.Integration;
using System;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;

namespace XQLite.AddIn
{
    public static class XqlFastTableReader
    {
        public sealed record Slice(string WorksheetName, string ListObjectName, int BodyTop, int BodyRows, int Cols, int BodyLeft);

        public static List<Slice> Collect()
        {
            var app = (Excel.Application)ExcelDnaUtil.Application; var list = new List<Slice>();
            foreach (Excel.Worksheet ws in app.Worksheets)
            {
                if (ws.ListObjects.Count == 0) continue;
                foreach (Excel.ListObject lo in ws.ListObjects)
                {
                    var body = lo.DataBodyRange; if (body == null) continue; var header = lo.HeaderRowRange; int cols = header.Columns.Count;
                    list.Add(new Slice(ws.Name, lo.Name, body.Row, body.Rows.Count, cols, body.Column));
                }
            }
            return list;
        }

        public static string[] ReadHeader(Slice s)
        {
            var app = (Excel.Application)ExcelDnaUtil.Application; var ws = (Excel.Worksheet)app.Worksheets[s.WorksheetName];
            var lo = ws.ListObjects[s.ListObjectName]; var header = lo.HeaderRowRange; int colCount = header.Columns.Count; var arr = (object[,])header.Value2;
            var headers = new string[colCount];
            for (int c = 1; c <= colCount; c++) headers[c - 1] = Convert.ToString(arr[1, c]) ?? $"C{c}";
            return headers;
        }

        public static IEnumerable<Dictionary<string, object?>> ReadBodyChunks(Slice s, string[] headers, int startRow1, int take)
        {
            var app = (Excel.Application)ExcelDnaUtil.Application; var ws = (Excel.Worksheet)app.Worksheets[s.WorksheetName];
            var lo = ws.ListObjects[s.ListObjectName]; var body = lo.DataBodyRange!; int colCount = headers.Length;
            int endRow1 = Math.Min(startRow1 + take - 1, s.BodyRows);
            var seg = body.Range[body.Cells[startRow1, 1], body.Cells[endRow1, colCount]]; var arr = (object[,])seg.Value2;
            for (int r = 1; r <= endRow1 - startRow1 + 1; r++)
            {
                var d = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
                for (int c = 1; c <= colCount; c++)
                {
                    var v = arr[r, c];
                    if (v is string ss) { ss = ss.Trim(); if (ss.Length == 0) v = null; else v = ss; }
                    d[headers[c - 1]] = v; // 숫자/날짜 등은 Value2 원본 유지
                }
                yield return d;
            }
        }
    }
}