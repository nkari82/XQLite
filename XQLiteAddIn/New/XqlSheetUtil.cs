// XqlSheetUtil.cs
using System;
using System.Collections.Generic;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;

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
                    if (string.IsNullOrEmpty(txt)) { XqlCommon.ReleaseCom(cell); break; }
                    XqlCommon.ReleaseCom(lastCell); lastCell = cell;
                }
                var lastCol = Math.Max(1, (lastCell?.Column as int?) ?? 1);
                var rg = sh.Range[sh.Cells[1, 1], sh.Cells[1, lastCol]];
                XqlCommon.ReleaseCom(lastCell);
                return rg;
            }
            catch { XqlCommon.ReleaseCom(lastCell); return (Excel.Range)sh.Cells[1, 1]; }
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
            finally { XqlCommon.ReleaseCom(header); }
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
                var used = ws.UsedRange; int lastRow = used.Row + used.Rows.Count - 1; XqlCommon.ReleaseCom(used);
                for (int r = firstDataRow; r <= lastRow; r++)
                {
                    Excel.Range? cell = null;
                    try
                    {
                        cell = (Excel.Range)ws.Cells[r, keyCol];
                        var v = cell.Value2;
                        if (EqualKey(v, key)) return r;
                    }
                    finally { XqlCommon.ReleaseCom(cell); }
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
                    finally { XqlCommon.ReleaseCom(cell); }
                }
            }
            finally { XqlCommon.ReleaseCom(header); }
        }

        // 안전 교차
        private static Excel.Range? IntersectSafe(Excel.Worksheet ws, Excel.Range a, Excel.Range b)
        {
            try { return ws.Application.Intersect(a, b); }
            catch { return null; }
        }

        /// <summary>워크시트 이름으로 찾기(정확 일치, Ordinal)</summary>
        public static Excel.Worksheet? FindWorksheet(Excel.Application app, string sheetName)
        {
            if (app is null || string.IsNullOrWhiteSpace(sheetName)) return null;

            foreach (Excel.Worksheet? ws in app.Worksheets)
            {
                try
                {
                    if (ws is null) continue;
                    if (string.Equals(ws.Name, sheetName, StringComparison.Ordinal))
                        return ws;
                }
                finally { XqlCommon.ReleaseCom(ws); }
            }
            return null;
        }

        /// <summary>워크시트 내 ListObject를 이름으로 찾기(정확 일치, Ordinal)</summary>
        public static Excel.ListObject? FindListObject(Excel.Worksheet ws, string listObjectName)
        {
            if (ws is null || string.IsNullOrWhiteSpace(listObjectName)) return null;

            foreach (Excel.ListObject? lo in ws.ListObjects)
            {
                try
                {
                    if (lo is null) continue;
                    if (string.Equals(lo.Name, listObjectName, StringComparison.Ordinal))
                        return lo;
                }
                finally { XqlCommon.ReleaseCom(lo); }
            }
            return null;
        }

        /// <summary>
        /// 주어진 Range를 포함(겹침)하는 ListObject 찾기.
        /// - 표의 데이터 바디나 헤더 어느 쪽이든 겹치면 그 표를 반환
        /// - 여러 표가 겹치면 첫 번째를 반환
        /// </summary>
        public static Excel.ListObject? FindListObjectContaining(Excel.Worksheet ws, Excel.Range rng)
        {
            if (ws is null || rng is null) return null;

            foreach (Excel.ListObject? lo in ws.ListObjects)
            {
                Excel.Range? header = null;
                Excel.Range? body = null;
                try
                {
                    if (lo is null) continue;

                    header = lo.HeaderRowRange;
                    body = lo.DataBodyRange;

                    // 헤더 또는 바디와 교차하면 포함으로 간주
                    if ((header != null && IntersectSafe(ws, rng, header) != null) ||
                        (body != null && IntersectSafe(ws, rng, body) != null))
                    {
                        return lo;
                    }
                }
                catch
                {
                    // ignore and continue
                }
                finally
                {
                    XqlCommon.ReleaseCom(header);
                    XqlCommon.ReleaseCom(body);
                    XqlCommon.ReleaseCom(lo);
                }
            }
            return null;
        }

        /// <summary>
        /// 테이블 논리명으로 ListObject를 찾습니다.
        /// - 허용 입력:
        ///   1) "ListObjectName"
        ///   2) "SheetName.ListObjectName"
        /// - 현재 워크시트 우선 검색 후, 필요 시 시트명 접두사가 주어지면 해당 시트로 이동해 재검색합니다.
        /// </summary>
        public static Excel.ListObject? FindListObjectByTable(Excel.Worksheet ws, string tableNameOrQualified)
        {
            if (ws is null || string.IsNullOrWhiteSpace(tableNameOrQualified)) return null;

            // 1) 분해: "Sheet.Table" → (sheet, loName), "Table" → (null, loName)
            string? sheetHint = null;
            string loName = tableNameOrQualified;
            var dot = tableNameOrQualified.IndexOf('.');
            if (dot >= 0)
            {
                sheetHint = tableNameOrQualified.Substring(0, dot);
                loName = tableNameOrQualified.Substring(dot + 1);
            }

            // 현재 시트에서 먼저 시도
            var found = FindListObject(ws, loName);
            if (found != null) return found;

            // 2) 헤더 텍스트(표시명)로도 한 번 더 시도
            found = FindListObjectByHeaderCaption(ws, loName);
            if (found != null) return found;

            // 3) sheetHint가 있으면 해당 시트로 이동해서 다시 검색
            if (sheetHint != null && sheetHint != "")
            {
                Excel.Worksheet? ws2 = null;
                try
                {
                    var app = ws.Application as Excel.Application;
                    ws2 = FindWorksheet(app!, sheetHint);
                    if (ws2 != null)
                    {
                        var f2 = FindListObject(ws2, loName);
                        if (f2 != null) return f2;

                        f2 = FindListObjectByHeaderCaption(ws2, loName);
                        if (f2 != null) return f2;
                    }
                }
                finally { XqlCommon.ReleaseCom(ws2); }
            }

            // 4) 마지막 시도: 현재 시트의 모든 표에서 "Sheet.ListObject" 정규화 비교
            foreach (Excel.ListObject? lo in ws.ListObjects)
            {
                try
                {
                    if (lo is null) continue;
                    var qualified = $"{ws.Name}.{lo.Name}";
                    if (string.Equals(qualified, tableNameOrQualified, StringComparison.Ordinal))
                        return lo;
                }
                finally { XqlCommon.ReleaseCom(lo); }
            }

            return null;
        }

        /// <summary>
        /// 헤더 캡션(표시 텍스트)로 ListObject 추정
        /// - 엑셀 표 이름과 별개로, UI 상단 헤더 텍스트가 테이블명인 경우를 완화 지원
        /// </summary>
        private static Excel.ListObject? FindListObjectByHeaderCaption(Excel.Worksheet ws, string caption)
        {
            foreach (Excel.ListObject? lo in ws.ListObjects)
            {
                Excel.Range? header = null;
                try
                {
                    if (lo is null) continue;
                    header = lo.HeaderRowRange;
                    if (header == null) continue;

                    var v = header.Value2 as object[,];
                    if (v == null) continue;

                    // 첫 셀 텍스트/전체 연결 텍스트 둘 다 비교
                    string first = Convert.ToString(v[1, 1]) ?? string.Empty;
                    if (string.Equals(first, caption, StringComparison.Ordinal))
                        return lo;

                    // "A|B|C" 처럼 헤더 전체를 합쳐서 비교하는 느슨한 추정 (과한 매칭 방지용 equals)
                    int cols = header.Columns.Count;
                    var joined = string.Empty;
                    for (int i = 1; i <= cols; i++)
                        joined += (i == 1 ? "" : "|") + (Convert.ToString(v[1, i]) ?? string.Empty);
                    if (string.Equals(joined, caption, StringComparison.Ordinal))
                        return lo;
                }
                catch { /* ignore */ }
                finally { XqlCommon.ReleaseCom(header); XqlCommon.ReleaseCom(lo); }
            }
            return null;
        }

        /* relative utils */

        // 직렬화
        public static string ColumnKey(string sheet, string table, int hRow, int hCol, int colOffset, string? hdrName = null)
            => hdrName is { Length: > 0 }
               ? $"col:{sheet}:{table}:H{hRow}C{hCol}:dx={colOffset}:hdr={Escape(hdrName)}"
               : $"col:{sheet}:{table}:H{hRow}C{hCol}:dx={colOffset}";

        public static string CellKey(string sheet, string table, int hRow, int hCol, int rowOffset, int colOffset)
            => $"cell:{sheet}:{table}:H{hRow}C{hCol}:dr={rowOffset}:dc={colOffset}";

        // 파싱
        public static bool TryParse(string key, out RelDesc d)
        {
            d = default;
            // 형식: col:Sheet:Table:H{r}C{c}:dx={dx}[:hdr=...]
            //     | cell:Sheet:Table:H{r}C{c}:dr={dr}:dc={dc}
            try
            {
                var parts = key.Split(':');
                if (parts.Length < 4) return false;
                var kind = parts[0];
                var sheet = parts[1];
                var table = parts[2];

                int hRow = 0, hCol = 0, dr = 0, dc = 0; string? hdr = null;

                foreach (var seg in parts.Skip(3))
                {
                    if (seg.StartsWith("H") && seg.Contains('C'))
                    {
                        var hc = seg.Substring(1).Split('C');
                        hRow = int.Parse(hc[0]);
                        hCol = int.Parse(hc[1]);
                    }
                    else if (seg.StartsWith("dx=")) dc = ParseIntOrDefault(seg, 3);      // 컬럼 offset
                    else if (seg.StartsWith("dr=")) dr = ParseIntOrDefault(seg, 3);      // 행 offset
                    else if (seg.StartsWith("dc=")) dc = ParseIntOrDefault(seg, 3);
                    else if (seg.StartsWith("hdr=")) hdr = Unescape(seg.Substring(4));
                }

                d = new RelDesc(kind, sheet, table, hRow, hCol, dr, dc, hdr);
                return true;
            }
            catch { return false; }
        }

        // 복원: 현재 Workbook에서 키를 실제 Range or Column Index로 변환
        public static bool TryResolve(Excel.Application app, RelDesc d, out Excel.Range? target, out int? targetHeaderCol, out Excel.ListObject? lo)
        {
            target = null; targetHeaderCol = null; lo = null;
            try
            {
                var ws = FindWorksheet(app, d.Sheet);
                if (ws == null) return false;

                lo = FindListObject(ws, d.Table);
                if (lo == null || lo.HeaderRowRange == null) return false;

                var hdr = lo.HeaderRowRange;
                int anchorRow = hdr.Row;
                int anchorCol = hdr.Column;

                if (anchorRow != d.AnchorRow || anchorCol != d.AnchorCol)
                {
                    // 앵커가 변했더라도 lo.HeaderRowRange가 현재 기준점 ⇒ 상대 오프셋만 적용하면 OK
                }

                if (d.Kind == "col")
                {
                    int col = anchorCol + d.ColOffset;
                    targetHeaderCol = col;
                    target = (Excel.Range)ws.Cells[anchorRow, col];
                    return true;
                }
                else if (d.Kind == "cell")
                {
                    int row = (anchorRow + 1) + d.RowOffset;   // 데이터 첫행 = header+1
                    int col = anchorCol + d.ColOffset;
                    target = (Excel.Range)ws.Cells[row, col];
                    return true;
                }
                return false;
            }
            catch { return false; }
        }

        internal readonly record struct RelDesc(string Kind, string Sheet, string Table, int AnchorRow, int AnchorCol, int RowOffset, int ColOffset, string? HeaderNameHint);

        private static string Escape(string s) => s.Replace("\\", "\\\\").Replace(":", "\\:");

        private static string Unescape(string s) => s.Replace("\\:", ":").Replace("\\\\", "\\");

        private static int ParseIntOrDefault(string s, int startIndex, int defaultValue = 0)
        {
            if (string.IsNullOrEmpty(s) || startIndex >= s.Length) return defaultValue;
            if (int.TryParse(s.Substring(startIndex), System.Globalization.NumberStyles.Integer,
                             System.Globalization.CultureInfo.InvariantCulture, out var v))
                return v;
            return defaultValue;
        }
    }
}
