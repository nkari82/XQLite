// XqlSheetUtil.cs
using System;
using System.Collections.Generic;
using System.Linq;
using ExcelDna.Integration;
using Excel = Microsoft.Office.Interop.Excel;
using static XQLite.AddIn.XqlCommon;

namespace XQLite.AddIn
{
    internal static class XqlSheetUtil
    {
        public static Excel.Range GetHeaderRange(Excel.Worksheet sh)
        {
            Excel.Range? lastCell = null;
            try
            {
                int col = 1;
                for (; col <= sh.UsedRange.Columns.Count; col++)
                {
                    var cell = (Excel.Range)sh.Cells[1, col];
                    var txt = (cell?.Value2 as string)?.Trim();
                    if (string.IsNullOrEmpty(txt)) { ReleaseCom(cell); break; }
                    ReleaseCom(lastCell); lastCell = cell;
                }
                var lastCol = Math.Max(1, (lastCell?.Column as int?) ?? 1);
                var rg = sh.Range[sh.Cells[1, 1], sh.Cells[1, lastCol]];
                ReleaseCom(lastCell);
                return rg;
            }
            catch { ReleaseCom(lastCell); return (Excel.Range)sh.Cells[1, 1]; }
        }

        public static (Excel.Range header, List<string> names) GetHeaderAndNames(Excel.Worksheet ws)
        {
            Excel.Range? header = null;
            var names = new List<string>();
            try
            {
                header = ws.Range[ws.Cells[1, 1], ws.Cells[1, ws.UsedRange.Columns.Count]];
                int cols = header.Columns.Count;
                for (int c = 1; c <= cols; c++)
                    names.Add(Convert.ToString(((Excel.Range)header.Cells[1, c]).Value2) ?? string.Empty);
                return (header, names.Select(s => s.Trim()).ToList());
            }
            catch { return ((Excel.Range)ws.Cells[1, 1], names); }
            finally { ReleaseCom(header); }
        }

        public static int FindKeyColumnIndex(List<string> headers, string keyName)
        {
            if (!string.IsNullOrWhiteSpace(keyName))
            {
                var idx = headers.FindIndex(h => string.Equals(h, keyName, StringComparison.Ordinal));
                if (idx >= 0) return idx + 1;
            }
            var id = headers.FindIndex(h => string.Equals(h, "id", StringComparison.OrdinalIgnoreCase));
            if (id >= 0) return id + 1;
            var key = headers.FindIndex(h => string.Equals(h, "key", StringComparison.OrdinalIgnoreCase));
            if (key >= 0) return key + 1;
            return 1;
        }

        public static int? FindRowByKey(Excel.Worksheet ws, int firstDataRow, int keyCol, object key)
        {
            try
            {
                var used = ws.UsedRange; int lastRow = used.Row + used.Rows.Count - 1; ReleaseCom(used);
                for (int r = firstDataRow; r <= lastRow; r++)
                {
                    Excel.Range? cell = null;
                    try
                    {
                        cell = (Excel.Range)ws.Cells[r, keyCol];
                        var v = cell.Value2;
                        if (EqualKey(v, key)) return r;
                    }
                    finally { ReleaseCom(cell); }
                }
            }
            catch { }
            return null;
        }

        public static bool EqualKey(object? a, object? b)
        {
            if (a is null && b is null) return true;
            if (a is null || b is null) return false;
            return string.Equals(Convert.ToString(a), Convert.ToString(b), StringComparison.Ordinal);
        }

        public static void SetHeaderTooltips(Excel.Worksheet sh, IReadOnlyDictionary<string, string> colToTip,
            Action<Excel.Range>? clear = null, Action<Excel.Range, string>? set = null)
        {
            var header = GetHeaderRange(sh);
            try
            {
                foreach (Excel.Range cell in header.Cells)
                {
                    try
                    {
                        var colName = (cell.Value2 as string)?.Trim();
                        if (string.IsNullOrEmpty(colName)) continue;
#pragma warning disable CS8604 // 가능한 null 참조 인수입니다.
                        if (!colToTip.TryGetValue(colName, out var tip)) continue;
#pragma warning restore CS8604 // 가능한 null 참조 인수입니다.
                        clear?.Invoke(cell); set?.Invoke(cell, tip);
                    }
                    finally { ReleaseCom(cell); }
                }
            }
            finally { ReleaseCom(header); }
        }
    }
}
