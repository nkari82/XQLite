// 시트의 헤더 표현
using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;
using System.Drawing;

#if false
namespace XQLite.AddIn
{
    internal static class XqlSheetMetaRegistry
    {
        private const string SheetMetaName = "_XQL_META";

        // 팔레트 & 스타일
        private static readonly int HeaderFill = Ole("#F6F7F9");
        private static readonly int FontColor = Ole("#2F3B52");
        private static readonly int BorderThin = Ole("#C9CED6");
        private static readonly int BorderInside = Ole("#D8DCE2");
        private static readonly int BorderBottom = Ole("#8893A3");
        private static readonly int ShadowTop = Ole("#E9EBF0");

        private const Excel.XlLineStyle LS = Excel.XlLineStyle.xlContinuous;

        public sealed class Meta { public int TopRow { get; init; } public int LeftCol { get; init; } public int ColCount { get; init; } }

        public static bool Exists(Excel.Worksheet ws) => FindAllMetaNames(ws).Count > 0;

        public static Meta? Get(Excel.Worksheet ws)
        {
            var names = FindAllMetaNames(ws);
            if (names.Count == 0) return null;
            foreach (var n in names)
            {
                try { var r = n.RefersToRange; if (r != null) return new Meta { TopRow = r.Row, LeftCol = r.Column, ColCount = r.Columns.Count }; }
                catch { }
            }
            return null;
        }

        /// 선택 행을 메타헤더로 설치 (freezePane 기본 false 권장)
        public static void CreateFromSelection(Excel.Worksheet ws, Excel.Range sel, bool freezePane = false)
        {
            if (ws is null || sel is null) throw new ArgumentNullException(nameof(sel));
            if (Exists(ws)) throw new InvalidOperationException("이 시트에는 이미 메타 헤더가 있습니다.");

            Excel.Range row;
            try { row = (Excel.Range)sel.Rows[1]; } catch { throw new InvalidOperationException("헤더로 사용할 셀 한 줄을 선택하세요."); }

            int top = row.Row, left = row.Column;
            int cnt; try { cnt = Math.Max(1, Convert.ToInt32(((Excel.Range)row.Columns).Count)); } catch { cnt = 1; }

            var header = ws.Range[ws.Cells[top, left], ws.Cells[top, left + cnt - 1]];
            TryCleanWorkbookScopeName(ws);

            string refersTo = "=" + header.get_Address(true, true, Excel.XlReferenceStyle.xlA1, true);
            ws.Names.Add(SheetMetaName, refersTo);
            if (!Exists(ws)) throw new InvalidOperationException("메타 헤더 이름 등록 실패");

            ApplyHeaderStyle(ws, header);

            if (freezePane)
            {
                try { var aw = ws.Application.ActiveWindow; aw.SplitRow = top; aw.SplitColumn = 0; aw.FreezePanes = true; } catch { }
            }
        }

        /// 복사/삭제/열 변경 후 호출: 폭 자동 갱신 + 스타일 재적용 (+ 컬럼검증 재적용)
        public static void RefreshHeaderBorders(Excel.Worksheet ws)
        {
            var meta = Get(ws);
            if (meta == null) return;

            var header = EnsureMetaRangeUpToDate(ws, meta);
            ApplyHeaderStyle(ws, header);

            // 컬럼 타입 검증 재적용
            XqlColumnTypeRegistry.ApplyAllForHeader(ws);
        }

        public static void Remove(Excel.Worksheet ws)
        {
            var names = FindAllMetaNames(ws);
            if (names.Count == 0) return;

            Excel.Range? header = null;
            foreach (var n in names) { try { header ??= n.RefersToRange as Excel.Range; } catch { } }

            if (header != null)
            {
                ClearHeaderStyle(ws, header);
                ClearRowBottomBorders(ws, header.Row);
                RemoveLeftoverLongLines(ws, header);
            }
            foreach (var n in names) { try { n.Delete(); } catch { } }
        }

        // ---------- internals ----------
        private static List<Excel.Name> FindAllMetaNames(Excel.Worksheet ws)
        {
            var list = new List<Excel.Name>();
            var wb = ws.Application.ActiveWorkbook;
            try { foreach (Excel.Name n in ws.Names) if ((n.NameLocal ?? n.Name) == SheetMetaName) list.Add(n); } catch { }
            try
            {
                foreach (Excel.Name n in wb.Names)
                {
                    var nm = n.NameLocal ?? n.Name;
                    if (!(nm.EndsWith("!" + SheetMetaName) || nm == SheetMetaName)) continue;
                    try { var r = n.RefersToRange; if (r != null && r.Worksheet?.Name == ws.Name) { list.Add(n); continue; } } catch { }
                    try
                    {
                        var s = n.RefersTo as string;
                        var m = !string.IsNullOrEmpty(s) ? Regex.Match(s!, @"=\s*'?([^'!]+)'?\s*!") : null;
                        if (m is { Success: true } && m.Groups[1].Value == ws.Name) list.Add(n);
                    }
                    catch { }
                }
            }
            catch { }
            return list;
        }

        private static void TryCleanWorkbookScopeName(Excel.Worksheet ws)
        {
            try
            {
                var wb = ws.Application.ActiveWorkbook;
                foreach (Excel.Name n in wb.Names)
                    if ((n.NameLocal ?? n.Name) == SheetMetaName) { try { n.Delete(); } catch { } }
            }
            catch { }
        }

        private static Excel.Range EnsureMetaRangeUpToDate(Excel.Worksheet ws, Meta meta)
        {
            int row = meta.TopRow, col = meta.LeftCol;
            int last = col;

            int usedCols; try { usedCols = Math.Max(ws.UsedRange?.Columns?.Count ?? 0, 1); } catch { usedCols = 64; }
            int maxC = Math.Min(16384, col + usedCols + 64);

            for (int c = col; c <= maxC; c++)
            {
                var cell = ws.Cells[row, c] as Excel.Range;
                if (cell == null) break;
                bool bold = false; int? fill = null;
                try { bold = Convert.ToBoolean(cell.Font.Bold); } catch { }
                try { fill = Convert.ToInt32(cell.Interior.Color); } catch { }
                if (bold && fill.HasValue && ColorsClose(fill.Value, HeaderFill)) last = c;
                else break;
            }

            int newCount = Math.Max(1, last - col + 1);
            if (newCount != meta.ColCount)
            {
                var newHeader = ws.Range[ws.Cells[row, col], ws.Cells[row, col + newCount - 1]];
                string refersTo = "=" + newHeader.get_Address(true, true, Excel.XlReferenceStyle.xlA1, true);
                foreach (var n in FindAllMetaNames(ws)) { try { n.RefersTo = refersTo; } catch { } }
                return newHeader;
            }
            return ws.Range[ws.Cells[row, col], ws.Cells[row, col + meta.ColCount - 1]];
        }

        private static void ApplyHeaderStyle(Excel.Worksheet ws, Excel.Range header)
        {
            ClearHeaderStyle(ws, header);

            // 배경 & 텍스트
            try
            {
                header.Interior.Color = HeaderFill;
                header.Font.Bold = true;
                header.Font.Color = FontColor;
                try { header.Font.Size = Math.Max(8, Convert.ToDouble(header.Font.Size) + 1); } catch { }
                header.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                header.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                header.WrapText = false;
                try { if (Convert.ToSingle(header.RowHeight) < 18f) header.RowHeight = 18f; } catch { }
            }
            catch { }

            // 상/좌 얇은 선
            TryBorder(header.Borders[Excel.XlBordersIndex.xlEdgeTop], LS, BorderThin, Excel.XlBorderWeight.xlThin);
            var first = header.Cells[1, 1] as Excel.Range ?? header;
            TryBorder(first.Borders[Excel.XlBordersIndex.xlEdgeLeft], LS, BorderThin, Excel.XlBorderWeight.xlThin);

            // 컬럼 구분
            TryBorder(header.Borders[Excel.XlBordersIndex.xlInsideVertical], LS, BorderInside, Excel.XlBorderWeight.xlHairline);

            // 우측 외곽(마지막 셀)
            try
            {
                var lastCell = header.Cells[1, header.Columns.Count] as Excel.Range ?? header;
                TryBorder(lastCell.Borders[Excel.XlBordersIndex.xlEdgeRight], LS, BorderThin, Excel.XlBorderWeight.xlThin);
            }
            catch { }

            // 하단 Medium
            TryBorder(header.Borders[Excel.XlBordersIndex.xlEdgeBottom], LS, BorderBottom, Excel.XlBorderWeight.xlMedium);

            // 다음 행 top 얇게(그림자 느낌)
            try
            {
                int rowBelow = header.Row + 1;
                var below = ws.Range[ws.Cells[rowBelow, header.Column], ws.Cells[rowBelow, header.Column + header.Columns.Count - 1]];
                TryBorder(below.Borders[Excel.XlBordersIndex.xlEdgeTop], LS, ShadowTop, Excel.XlBorderWeight.xlHairline);
            }
            catch { }
        }

        private static void ClearHeaderStyle(Excel.Worksheet ws, Excel.Range header)
        {
            void Off(Excel.XlBordersIndex i) { try { header.Borders[i].LineStyle = Excel.XlLineStyle.xlLineStyleNone; } catch { } }
            Off(Excel.XlBordersIndex.xlEdgeTop);
            Off(Excel.XlBordersIndex.xlEdgeBottom);
            Off(Excel.XlBordersIndex.xlInsideVertical);
            try { var first = header.Cells[1, 1] as Excel.Range; first?.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlLineStyleNone; } catch { }
            try { var last = header.Cells[1, header.Columns.Count] as Excel.Range; last?.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlLineStyleNone; } catch { }

            try { header.Interior.Pattern = Excel.XlPattern.xlPatternNone; header.Interior.ColorIndex = Excel.XlColorIndex.xlColorIndexNone; } catch { }
            try { header.Font.Bold = false; header.Font.ColorIndex = Excel.XlColorIndex.xlColorIndexAutomatic; } catch { }
            try { header.HorizontalAlignment = Excel.XlHAlign.xlHAlignGeneral; header.VerticalAlignment = Excel.XlVAlign.xlVAlignBottom; } catch { }

            // 아래 행 top 라인 제거
            try
            {
                int rowBelow = header.Row + 1;
                var below = ws.Range[ws.Cells[rowBelow, header.Column], ws.Cells[rowBelow, header.Column + header.Columns.Count - 1]];
                below.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlLineStyleNone;
            }
            catch { }
        }

        private static void ClearRowBottomBorders(Excel.Worksheet ws, int row)
        {
            try
            {
                int lastCol; try { lastCol = Math.Max(ws.UsedRange?.Columns?.Count ?? 0, 1); } catch { lastCol = 64; }
                lastCol = Math.Min(lastCol + 8, 16384);
                for (int c = 1; c <= lastCol; c++)
                    try { (ws.Cells[row, c] as Excel.Range)!.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlLineStyleNone; } catch { }
                try { (ws.Rows[row] as Excel.Range)!.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlLineStyleNone; } catch { }
            }
            catch { }
        }

        private static void RemoveLeftoverLongLines(Excel.Worksheet ws, Excel.Range header)
        {
            try
            {
                var toDel = new List<string>();
                double top = 0, w = 0, hW = 0; try { top = Convert.ToDouble(header.Top); } catch { }
                try { hW = Convert.ToDouble(header.Width); } catch { }

                foreach (Excel.Shape s in ws.Shapes)
                {
                    try
                    {
                        if (s.Type != Microsoft.Office.Core.MsoShapeType.msoLine) continue;
                        try { w = Convert.ToDouble(s.Width); } catch { w = 0; }
                        double st = 0; try { st = Convert.ToDouble(s.Top); } catch { }

                        bool near = Math.Abs(st - top) <= 8.0;
                        bool longLine = (hW > 0) ? w > Math.Max(300.0, hW * 1.2) : w > 1500.0;
                        if (near && longLine)
                        {
                            var nm = s.Name ?? $"DEL_TMP_{Guid.NewGuid():N}";
                            try { s.Name = nm; } catch { }
                            toDel.Add(nm);
                        }
                    }
                    catch { }
                }
                for (int i = toDel.Count - 1; i >= 0; i--) { try { ws.Shapes.Item(toDel[i]).Delete(); } catch { } }
            }
            catch { }
        }

        private static int Ole(string hex) => ColorTranslator.ToOle(ColorTranslator.FromHtml(hex));
        private static bool ColorsClose(int a, int b, int tol = 6)
        {
            var A = ColorTranslator.FromOle(a); var B = ColorTranslator.FromOle(b);
            return Math.Abs(A.R - B.R) <= tol && Math.Abs(A.G - B.G) <= tol && Math.Abs(A.B - B.B) <= tol;
        }
        private static void TryBorder(Excel.Border b, Excel.XlLineStyle ls, int color, Excel.XlBorderWeight w)
        {
            try { b.LineStyle = ls; b.Color = color; b.Weight = w; } catch { }
        }

        /// <summary>현재 시트의 메타 헤더 Range를 반환(없으면 null).</summary>
        public static Excel.Range? GetHeaderRange(Excel.Worksheet ws)
        {
            if (ws == null) return null;
            try
            {
                var names = FindAllMetaNames(ws);
                foreach (var n in names)
                {
                    try
                    {
                        var r = n.RefersToRange;
                        if (r != null) return r;
                    }
                    catch { /* ignore */ }
                }
            }
            catch { }
            return null;
        }

        // 메타 타입 딕셔너리 가져오기
        private static Dictionary<string, string> GetMetaTypes(Excel.Worksheet ws)
        {
            // 예: { "Name" : "string", "Level": "int", "Config": "json" }
            // SheetMetaRegistry.Get(ws) 에서 헤더행 읽고, 각 셀에 저장된 타입 읽어오기
            // 임시: Comment에서 읽는 방식
            var meta = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var header = XqlSheetMetaRegistry.GetHeaderRange(ws);
            if (header != null)
            {
                foreach (Excel.Range cell in header.Cells)
                {
                    string name = cell.Value2?.ToString() ?? "";
                    string type = cell.Comment?.Text() ?? "string";
                    if (!string.IsNullOrEmpty(name))
                        meta[name] = type;
                }
            }
            return meta;
        }

        // 셀 값 검증/변환
        private static object? NormalizeByType(object? v, string type)
        {
            if (v == null) return null;
            string s = v.ToString()?.Trim() ?? "";

            try
            {
                return type.ToLowerInvariant() switch
                {
                    "int" => (long)Convert.ToDouble(s),
                    "double" => Convert.ToDouble(s),
                    "bool" => bool.Parse(s),
                    "json" => XqlJson.Deserialize<object>(s),
                    _ => s // default string
                };
            }
            catch
            {
                // 변환 실패 → 그대로 문자열 반환 (또는 null 처리)
                return s;
            }
        }

        // 헤더에 타입 툴팁 표시
        private static void ApplyHeaderTooltips(Excel.Worksheet ws, Excel.Range header, Dictionary<string, string> meta)
        {
            foreach (Excel.Range cell in header.Cells)
            {
                string name = cell.Value2?.ToString() ?? "";
                if (string.IsNullOrEmpty(name)) continue;

                if (meta.TryGetValue(name, out string type))
                {
                    try
                    {
                        cell.ClearComments();
                        cell.AddComment($"Type: {type}");
                    }
                    catch { }
                }
            }
        }

    }
}
#endif