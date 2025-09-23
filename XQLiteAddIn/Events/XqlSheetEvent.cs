// XqlSheetEvents.cs
using System;
using System.Collections.Generic;
using System.Linq;
using ExcelDna.Integration;
using Excel = Microsoft.Office.Interop.Excel;

#if false
namespace XQLite.AddIn
{
    internal static class XqlSheetEvents
    {
        private static Excel.Application? _app;
        private static bool _hooked;

        internal static void Hook()
        {
            if (_hooked) return;
            _app = (Excel.Application)ExcelDnaUtil.Application;
            _app.SheetChange += App_SheetChange;
            _hooked = true;
            XqlLog.Info("SheetEvents hooked");
        }

        internal static void Unhook()
        {
            if (!_hooked || _app is null) 
                return;

            try 
            { 
                _app.SheetChange -= App_SheetChange; 
            } 
            catch { }
            _hooked = false;
            XqlLog.Info("SheetEvents unhooked");
        }

        private static void App_SheetChange(object Sh, Excel.Range Target)
        {
            try
            {
                if (Target == null) return;
                var ws = Sh as Excel.Worksheet ?? Target.Worksheet as Excel.Worksheet;
                if (ws == null) return;

                // 테이블 바디가 아니면 무시
                if (Target.ListObject is not Excel.ListObject lo) 
                    return;

                var header = lo.HeaderRowRange;
                var body = lo.DataBodyRange;
                if (body == null)
                    return;               // 빈 표

                if (IsIntersect(Target, header))
                    return; // 헤더 수정 무시

                // 락 검사 → Undo
                if (IsLockedEdit(ws, lo, Target))
                {
                    ((Excel.Application)ExcelDnaUtil.Application).Undo();
                    return;
                }

                // 영향받은 행들만 업서트 큐에 적재
                EnqueueRangeChange(ws, lo, Target);
            }
            catch (Exception ex)
            {
                XqlLog.Warn("SheetChange handler error: " + ex.Message);
            }
        }

        private static void EnqueueRangeChange(Excel.Worksheet ws, Excel.ListObject lo, Excel.Range target)
        {
            try
            {
                var body = lo.DataBodyRange!;
                var header = lo.HeaderRowRange!;
                int colCount = header.Columns.Count;

                var inter = IntersectSafe(target, body);
                if (inter == null) return;

                var rowsIdx = GetRelativeRowIndexesWithinBody(body, inter);
                if (rowsIdx.Count == 0) return;

                var headers = ReadHeaders(header, colCount);
                string tableName = XqlTableNameMap.Map(lo.Name, ws.Name);

                foreach (var relRow in rowsIdx)
                {
                    var rowDict = ReadRowDict(body, headers, relRow, colCount);
                    if (IsEmptyRow(rowDict)) continue;
                    XqlUpsert.Enqueue(tableName, rowDict);
                }
            }
            catch (Exception ex)
            {
                XqlLog.Warn("EnqueueRangeChange error: " + ex.Message);
            }
        }

        private static bool IsLockedEdit(Excel.Worksheet ws, Excel.ListObject lo, Excel.Range target)
        {
            try
            {
                var header = lo.HeaderRowRange!;
                int firstCol = target.Column;
                int lastCol = target.Column + target.Columns.Count - 1;

                for (int col = firstCol; col <= lastCol; col++)
                {
                    int colIndex = col - header.Column + 1;
                    if (colIndex < 1 || colIndex > header.Columns.Count) 
                        continue;

                    var headerArr = (object[,])header.Value2;
                    string colName = Convert.ToString(headerArr[1, colIndex]) ?? "";
                    string tableName = XqlTableNameMap.Map(lo.Name, ws.Name);

                    if (XqlLockService.IsLockedColumn(tableName, colName))
                    {
                        XqlLog.Warn($"Blocked by column lock: {tableName}.{colName}");
                        return true;
                    }
                }

                Excel.Range firstCell = (Excel.Range)target.Cells[1, 1];
                string addr = firstCell.Address[false, false];
                if (XqlLockService.IsLockedCell(ws.Name, addr))
                {
                    XqlLog.Warn($"Blocked by cell lock: {ws.Name}!{addr}");
                    return true;
                }

                return false;
            }
            catch { return false; }
        }

        private static bool IsIntersect(Excel.Range a, Excel.Range b)
        {
            try { return a.Worksheet.Application.Intersect(a, b) != null; }
            catch { return false; }
        }
        private static Excel.Range? IntersectSafe(Excel.Range a, Excel.Range b)
        {
            try { return a.Worksheet.Application.Intersect(a, b); }
            catch { return null; }
        }

        private static List<int> GetRelativeRowIndexesWithinBody(Excel.Range body, Excel.Range inter)
        {
            var list = new List<int>();
            int bodyTop = body.Row;
            int interTop = inter.Row;
            int interBottom = inter.Row + inter.Rows.Count - 1;

            for (int r = interTop; r <= interBottom; r++)
            {
                int rel = r - bodyTop + 1;
                if (rel >= 1 && rel <= body.Rows.Count) list.Add(rel);
            }
            return list.Distinct().ToList();
        }

        private static string[] ReadHeaders(Excel.Range header, int colCount)
        {
            var arr = (object[,])header.Value2;
            var headers = new string[colCount];
            for (int c = 1; c <= colCount; c++)
                headers[c - 1] = Convert.ToString(arr[1, c]) ?? $"C{c}";
            return headers;
        }

        private static Dictionary<string, object?> ReadRowDict(Excel.Range body, string[] headers, int relRow1, int colCount)
        {
            var cell = body.Cells[relRow1, 1];
            var rowRange = body.Range[cell, body.Cells[relRow1, colCount]];
            var arr = (object[,])rowRange.Value2;

            var d = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
            for (int c = 1; c <= colCount; c++)
            {
                d[headers[c - 1]] = NormalizeCellValue(arr[1, c]);
            }
            return d;
        }

        private static bool IsEmptyRow(Dictionary<string, object?> row)
        {
            foreach (var kv in row)
            {
                if (kv.Value is string s && s.Length > 0) return false;
                if (kv.Value != null && kv.Value is not string) return false;
            }
            return true;
        }

        private static object? NormalizeCellValue(object? v)
        {
            if (v == null) return null;

            if (v is double d)
            {
                // Excel 수치는 double → 정수 여부 판별
                if (Math.Abs(d % 1) < 1e-9) return (long)d;
                return d;
            }

            if (v is string s)
            {
                s = s.Trim();
                if (s.Length == 0) return null;

                // bool
                if (bool.TryParse(s, out var b)) return b;

                // number as string
                if (double.TryParse(s, out var dn)) return dn;

                // JSON?
                if ((s.StartsWith("{") && s.EndsWith("}")) || (s.StartsWith("[") && s.EndsWith("]")))
                {
                    try { return XqlJson.Deserialize<object>(s); }
                    catch { return s; }
                }

                return s;
            }

            return v; // fallback
        }
    }
}
#endif