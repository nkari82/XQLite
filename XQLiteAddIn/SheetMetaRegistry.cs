// SheetMetaRegistry.cs
using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;
using System.Drawing;

namespace XQLite.AddIn
{
    /// <summary>
    /// Sheet meta header 관리 (Borders 기반, 도형 미사용)
    /// - CreateFromSelection(ws, sel, freezePane=false)
    /// - Remove(ws)
    /// - RefreshHeaderBorders(ws)
    /// 
    /// 특징:
    /// - 헤더 배경은 연한 그레이로 채우고,
    /// - 상단 얇은 선, 컬럼 구분 얇은 세로선, 하단 굵은 선을 적용.
    /// - 제거 시 테두리/채우기/이전 도형 잔여물 모두 정리.
    /// - Initialize/이벤트 훅 없음(원하면 별도 추가).
    /// </summary>
    internal static class SheetMetaRegistry
    {
        private const string SheetMetaName = "_XQL_META";

        // 스타일 (변경 가능)
        private static readonly int ColorThin = ColorTranslator.ToOle(Color.FromArgb(170, 170, 170)); // 연한 그레이
        private static readonly int ColorBottom = ColorTranslator.ToOle(Color.FromArgb(100, 100, 100)); // 진한 그레이
        private static readonly int HeaderFillColor = ColorTranslator.ToOle(Color.FromArgb(242, 242, 242)); // #F2F2F2

        private const Excel.XlLineStyle ThinLineStyle = Excel.XlLineStyle.xlContinuous;
        private const Excel.XlLineStyle ThickLineStyle = Excel.XlLineStyle.xlContinuous;

        public sealed class Meta
        {
            public int TopRow { get; init; }
            public int LeftCol { get; init; }
            public int ColCount { get; init; }
        }

        // ---------------- Public API ----------------

        public static bool Exists(Excel.Worksheet ws) => FindAllMetaNames(ws).Count > 0;

        public static Meta? Get(Excel.Worksheet ws)
        {
            var names = FindAllMetaNames(ws);
            if (names.Count == 0) return null;
            foreach (var n in names)
            {
                try
                {
                    var r = n.RefersToRange;
                    if (r != null) return new Meta { TopRow = r.Row, LeftCol = r.Column, ColCount = r.Columns.Count };
                }
                catch { }
            }
            return null;
        }

        /// <summary>
        /// 선택한 한 줄을 메타 헤더로 등록.
        /// freezePane 기본 false (Freeze Panses가 화면 분할선을 길게 보이게 함).
        /// </summary>
        public static void CreateFromSelection(Excel.Worksheet ws, Excel.Range sel, bool freezePane = false)
        {
            if (ws is null || sel is null) throw new ArgumentNullException(nameof(sel));
            if (Exists(ws)) throw new InvalidOperationException("이 시트에는 이미 메타 헤더가 있습니다.");

            Excel.Range selRow;
            try { selRow = (Excel.Range)sel.Rows[1]; }
            catch { throw new InvalidOperationException("헤더로 사용할 셀 한 줄을 선택한 뒤 다시 실행하세요."); }

            int topRow = selRow.Row;
            int leftCol = selRow.Column;
            int colCount;
            try { object cnt = ((Excel.Range)selRow.Columns).Count; colCount = Math.Max(1, Convert.ToInt32(cnt)); } catch { colCount = 1; }

            var header = ws.Range[ws.Cells[topRow, leftCol], ws.Cells[topRow, leftCol + Math.Max(1, colCount) - 1]];

            TryCleanWorkbookScopeName(ws);

            // Use address string for compatibility
            string refersTo = "=" + header.get_Address(RowAbsolute: true, ColumnAbsolute: true, ReferenceStyle: Excel.XlReferenceStyle.xlA1, External: true);
            ws.Names.Add(SheetMetaName, refersTo);

            if (!Exists(ws)) throw new InvalidOperationException("메타 헤더 이름 등록에 실패했습니다.");

            // Apply borders and header fill
            ApplyHeaderBorders(ws, header);

            // Freeze optional; default false to avoid long split line issues
            if (freezePane)
            {
                try
                {
                    var app = ws.Application;
                    app.ActiveWindow.SplitRow = topRow;
                    app.ActiveWindow.SplitColumn = 0;
                    app.ActiveWindow.FreezePanes = true;
                }
                catch { }
            }

            // Protect shapes optionally (borders don't require protection; harmless)
            TryProtectSheetForShapes(ws);
        }

        /// <summary>
        /// 메타 제거: borders/interior 제거, 이전 도형 잔여물 강제 제거, 메타 이름 삭제
        /// </summary>
        public static void Remove(Excel.Worksheet ws)
        {
            var names = FindAllMetaNames(ws);
            if (names.Count == 0) return;

            Excel.Range? header = null;
            foreach (var n in names) { try { header ??= n.RefersToRange as Excel.Range; } catch { } }

            if (header != null)
            {
                bool didUnprotect = false;
                try
                {
                    try { ws.Unprotect(Type.Missing); didUnprotect = true; } catch { /* ignore */ }

                    // remove borders & fill
                    RemoveHeaderBorders(ws, header);

                    // clear entire row bottom borders more aggressively
                    try { ClearRowBottomBorders(ws, header.Row); } catch { }

                    // remove leftover shapes if any (aggressive)
                    RemoveLeftoverMetaShapesAggressive(ws, header);
                }
                finally
                {
                    if (didUnprotect)
                    {
                        try { ws.Protect(DrawingObjects: true, Contents: false, Scenarios: false, UserInterfaceOnly: true); } catch { }
                    }
                }
            }

            foreach (var n in names) { try { n.Delete(); } catch { } }
        }

        /// <summary>
        /// 수동으로 헤더 테두리를 재적용. (복사·삭제·열 너비 변경 시 호출)
        /// </summary>
        public static void RefreshHeaderBorders(Excel.Worksheet ws)
        {
            try
            {
                var meta = Get(ws);
                if (meta == null) return;
                var header = ws.Range[ws.Cells[meta.TopRow, meta.LeftCol], ws.Cells[meta.TopRow, meta.LeftCol + Math.Max(1, meta.ColCount) - 1]];
                ApplyHeaderBorders(ws, header);
            }
            catch { /* best-effort */ }
        }

        // ---------------- internal helpers ----------------

        private static List<Excel.Name> FindAllMetaNames(Excel.Worksheet ws)
        {
            var list = new List<Excel.Name>();
            var wb = ws.Application.ActiveWorkbook;

            try
            {
                foreach (Excel.Name n in ws.Names)
                {
                    var nm = n.NameLocal ?? n.Name;
                    if (string.Equals(nm, SheetMetaName, StringComparison.Ordinal)) list.Add(n);
                }
            }
            catch { }

            try
            {
                foreach (Excel.Name n in wb.Names)
                {
                    var nm = n.NameLocal ?? n.Name;
                    if (!nm.EndsWith("!" + SheetMetaName, StringComparison.Ordinal) && !string.Equals(nm, SheetMetaName, StringComparison.Ordinal)) continue;

                    try
                    {
                        var r = n.RefersToRange;
                        if (r != null && r.Worksheet != null && r.Worksheet.Name == ws.Name) { list.Add(n); continue; }
                    }
                    catch { }

                    try
                    {
                        var refersTo = n.RefersTo as string;
                        if (!string.IsNullOrEmpty(refersTo))
                        {
                            var m = Regex.Match(refersTo, @"=\s*'?([^'!]+)'?\s*!");
                            if (m.Success)
                            {
                                var sheetInRef = m.Groups[1].Value;
                                if (string.Equals(sheetInRef, ws.Name, StringComparison.Ordinal)) list.Add(n);
                            }
                        }
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
                {
                    var nm = n.NameLocal ?? n.Name;
                    if (string.Equals(nm, SheetMetaName, StringComparison.Ordinal))
                    {
                        try { n.Delete(); } catch { }
                    }
                }
            }
            catch { }
        }

        // ---------------- Borders drawing ----------------

        private static void ApplyHeaderBorders(Excel.Worksheet ws, Excel.Range header)
        {
            try { RemoveHeaderBorders(ws, header); } catch { }

            // header fill (gray)
            try { header.Interior.Color = HeaderFillColor; } catch { }

            // top thin border across header
            try
            {
                var bTop = header.Borders[Excel.XlBordersIndex.xlEdgeTop];
                bTop.LineStyle = ThinLineStyle;
                bTop.Color = ColorThin;
                bTop.Weight = Excel.XlBorderWeight.xlThin;
            }
            catch { }

            // leftmost border (first cell)
            try
            {
                var firstCell = header.Cells[1, 1] as Excel.Range ?? header;
                var bLeft = firstCell.Borders[Excel.XlBordersIndex.xlEdgeLeft];
                bLeft.LineStyle = ThinLineStyle;
                bLeft.Color = ColorThin;
                bLeft.Weight = Excel.XlBorderWeight.xlThin;
            }
            catch { }

            // per-column right edges
            int colCount = 1;
            try { colCount = Math.Max(1, Convert.ToInt32(header.Columns.Count)); } catch { colCount = 1; }
            for (int i = 1; i <= colCount; i++)
            {
                try
                {
                    var cell = header.Cells[1, i] as Excel.Range;
                    if (cell == null) continue;
                    var bRight = cell.Borders[Excel.XlBordersIndex.xlEdgeRight];
                    bRight.LineStyle = ThinLineStyle;
                    bRight.Color = ColorThin;
                    bRight.Weight = Excel.XlBorderWeight.xlThin;
                }
                catch { }
            }

            // bottom thick edge (header bottom)
            try
            {
                var bBottom = header.Borders[Excel.XlBordersIndex.xlEdgeBottom];
                bBottom.LineStyle = ThickLineStyle;
                bBottom.Color = ColorBottom;
                try { bBottom.Weight = Excel.XlBorderWeight.xlThick; } catch { bBottom.Weight = Excel.XlBorderWeight.xlMedium; }
            }
            catch { }

            // try protect sheet shapes (harmless for borders)
            TryProtectSheetForShapes(ws);
        }

        private static void RemoveHeaderBorders(Excel.Worksheet ws, Excel.Range header)
        {
            try
            {
                // top
                try { var bTop = header.Borders[Excel.XlBordersIndex.xlEdgeTop]; bTop.LineStyle = Excel.XlLineStyle.xlLineStyleNone; } catch { }

                // leftmost
                try
                {
                    var firstCell = header.Cells[1, 1] as Excel.Range;
                    if (firstCell != null) { var bLeft = firstCell.Borders[Excel.XlBordersIndex.xlEdgeLeft]; bLeft.LineStyle = Excel.XlLineStyle.xlLineStyleNone; }
                }
                catch { }

                // per-column right
                int colCount = 1;
                try { colCount = Math.Max(1, Convert.ToInt32(header.Columns.Count)); } catch { colCount = 1; }
                for (int i = 1; i <= colCount; i++)
                {
                    try
                    {
                        var cell = header.Cells[1, i] as Excel.Range;
                        if (cell == null) continue;
                        var bRight = cell.Borders[Excel.XlBordersIndex.xlEdgeRight];
                        bRight.LineStyle = Excel.XlLineStyle.xlLineStyleNone;
                    }
                    catch { }
                }

                // bottom
                try { var bBottom = header.Borders[Excel.XlBordersIndex.xlEdgeBottom]; bBottom.LineStyle = Excel.XlLineStyle.xlLineStyleNone; } catch { }

                // header fill clear
                try
                {
                    header.Interior.Pattern = Excel.XlPattern.xlPatternNone;
                    try { header.Interior.ColorIndex = Excel.XlColorIndex.xlColorIndexNone; } catch { }
                }
                catch { }
            }
            catch { /* best-effort */ }
        }

        // Clear bottom borders across the row more aggressively (removes long horizontal borders)
        private static void ClearRowBottomBorders(Excel.Worksheet ws, int row)
        {
            try
            {
                int lastCol = 0;
                try
                {
                    var ur = ws.UsedRange;
                    if (ur != null) lastCol = Math.Max(ur.Columns.Count, 1);
                }
                catch { lastCol = 0; }

                if (lastCol <= 1)
                {
                    try { lastCol = ws.Columns.Count; } catch { lastCol = 256; }
                }

                int maxColToClear = Math.Min(lastCol + 5, Math.Min(16384, lastCol + 50));
                for (int c = 1; c <= maxColToClear; c++)
                {
                    try
                    {
                        var cell = ws.Cells[row, c] as Excel.Range;
                        if (cell == null) continue;
                        var b = cell.Borders[Excel.XlBordersIndex.xlEdgeBottom];
                        b.LineStyle = Excel.XlLineStyle.xlLineStyleNone;
                    }
                    catch { }
                }

                try
                {
                    var rowRange = ws.Rows[row] as Excel.Range;
                    if (rowRange != null)
                    {
                        var br = rowRange.Borders[Excel.XlBordersIndex.xlEdgeBottom];
                        br.LineStyle = Excel.XlLineStyle.xlLineStyleNone;
                    }
                }
                catch { }
            }
            catch { }
        }

        // Aggressive leftover shape cleanup (for previously shape-based implementations)
        private static void RemoveLeftoverMetaShapesAggressive(Excel.Worksheet ws, Excel.Range? header = null)
        {
            try
            {
                bool didUnprotect = false;
                try { ws.Unprotect(Type.Missing); didUnprotect = true; } catch { }

                var toDelete = new List<string>();
                double headerTop = double.NaN, headerWidth = double.NaN;
                if (header != null)
                {
                    try { headerTop = Convert.ToDouble(header.Top); } catch { headerTop = double.NaN; }
                    try { headerWidth = Convert.ToDouble(header.Width); } catch { headerWidth = double.NaN; }
                }

                foreach (Excel.Shape s in ws.Shapes)
                {
                    try
                    {
                        var nm = s.Name ?? "";
                        if (nm.StartsWith("_XQL_META_LINE_", StringComparison.Ordinal))
                        {
                            toDelete.Add(nm);
                            continue;
                        }

                        // candidate: line-like and very long near header
                        if (s.Type != Microsoft.Office.Core.MsoShapeType.msoLine 
                            /*&& s.Type != Microsoft.Office.Core.MsoShapeType.msoCurve*/) 
                            continue;

                        double top = 0, w = 0;
                        try { top = Convert.ToDouble(s.Top); } catch { }
                        try { w = Convert.ToDouble(s.Width); } catch { }

                        bool nearHeader = !double.IsNaN(headerTop) && Math.Abs(top - headerTop) <= 8.0;
                        bool veryLong = !double.IsNaN(headerWidth) ? (w > Math.Max(300.0, headerWidth * 1.2)) : (w > 1500.0);

                        if (nearHeader && veryLong)
                        {
                            if (!string.IsNullOrEmpty(nm)) toDelete.Add(nm);
                            else
                            {
                                try { s.Name = $"DEL_TMP_{Guid.NewGuid():N}"; toDelete.Add(s.Name); } catch { }
                            }
                        }
                    }
                    catch { }
                }

                for (int i = toDelete.Count - 1; i >= 0; i--)
                {
                    try { ws.Shapes.Item(toDelete[i]).Delete(); } catch { }
                }

                if (didUnprotect)
                {
                    try { ws.Protect(DrawingObjects: true, Contents: false, Scenarios: false, UserInterfaceOnly: true); } catch { }
                }
            }
            catch { }
        }

        private static void TryProtectSheetForShapes(Excel.Worksheet ws)
        {
            try { ws.Protect(DrawingObjects: true, Contents: false, Scenarios: false, UserInterfaceOnly: true); } catch { try { ws.Protect(Type.Missing, Type.Missing, false, false, true); } catch { } }
        }
    }
}
