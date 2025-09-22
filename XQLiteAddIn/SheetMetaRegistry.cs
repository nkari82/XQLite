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
    /// - RefreshHeaderBorders(ws) : 스타일/테두리 재적용 + 폭 자동 갱신
    /// - Remove(ws) : 테두리/채우기/잔여 도형 정리 + 이름 삭제
    ///
    /// 특징:
    /// - 헤더 복사/삭제 후 Refresh만 누르면 폭 자동 인식(헤더 시그니처 기반 스캔)
    /// - 헤더 스타일: 연한 그레이 배경, Bold(+1pt), 가운데 정렬, 얇은 상/좌/우, InsideVertical, 하단 Medium
    /// </summary>
    internal static class SheetMetaRegistry
    {
        private const string SheetMetaName = "_XQL_META";

        // 색/스타일
        private static readonly int ColorThin = ColorTranslator.ToOle(Color.FromArgb(190, 196, 205)); // 부드러운 회색
        private static readonly int ColorBottom = ColorTranslator.ToOle(Color.FromArgb(120, 130, 145)); // 약간 진함
        private static readonly int HeaderFillColor = ColorTranslator.ToOle(Color.FromArgb(244, 245, 247)); // #F4F5F7

        private const Excel.XlLineStyle ThinLineStyle = Excel.XlLineStyle.xlContinuous;
        private const Excel.XlLineStyle ThickLineStyle = Excel.XlLineStyle.xlContinuous;

        // 메타 DTO
        public sealed class Meta
        {
            public int TopRow { get; init; }
            public int LeftCol { get; init; }
            public int ColCount { get; init; }
        }

        // ---------- Public API ----------

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
        /// 선택한 한 줄을 메타 헤더로 등록 (freezePane 기본 false 권장)
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
            try { object cnt = ((Excel.Range)selRow.Columns).Count; colCount = Math.Max(1, Convert.ToInt32(cnt)); }
            catch { colCount = 1; }

            var header = ws.Range[ws.Cells[topRow, leftCol], ws.Cells[topRow, leftCol + colCount - 1]];

            TryCleanWorkbookScopeName(ws);

            string refersTo = "=" + header.get_Address(true, true, Excel.XlReferenceStyle.xlA1, true);
            ws.Names.Add(SheetMetaName, refersTo);

            if (!Exists(ws)) throw new InvalidOperationException("메타 헤더 이름 등록에 실패했습니다.");

            // 스타일 적용
            ApplyHeaderBorders(ws, header);

            // Freeze 선택(기본 false)
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
        }

        /// <summary>
        /// 복사/삭제/열 너비 변경 후 눌러주세요.
        /// - 현재 시그니처(배경+Bold) 기반으로 우측 연속 폭을 **자동 재탐색**해 메타 이름의 RefersTo 업데이트
        /// - 그 후 스타일/테두리 재적용
        /// </summary>
        public static void RefreshHeaderBorders(Excel.Worksheet ws)
        {
            var meta = Get(ws);
            if (meta == null) return;

            // 1) 현재 폭 재탐색 → 이름 RefersTo 업데이트
            var header = EnsureMetaRangeUpToDate(ws, meta);

            // 2) 스타일 재적용
            ApplyHeaderBorders(ws, header);
        }

        /// <summary>메타 제거: 테두리/채우기/잔여 도형 정리 + 이름 삭제</summary>
        public static void Remove(Excel.Worksheet ws)
        {
            var names = FindAllMetaNames(ws);
            if (names.Count == 0) return;

            Excel.Range? header = null;
            foreach (var n in names) { try { header ??= n.RefersToRange as Excel.Range; } catch { } }

            if (header != null)
            {
                // 테두리/채우기 제거
                RemoveHeaderBorders(ws, header);

                // 혹시 남아있을 라인 도형(이전 버전 잔재) 제거
                RemoveLeftoverLineShapes(ws, header);

                // 긴 하단선이 전역으로 남아있을 가능성 방지
                ClearRowBottomBorders(ws, header.Row);
            }

            foreach (var n in names) { try { n.Delete(); } catch { } }
        }

        // ---------- internal helpers ----------

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

        /// <summary>
        /// 현재 메타 헤더의 **실제 가로 폭**을 “헤더 시그니처”로 재탐색하여,
        /// _XQL_META 이름의 RefersTo를 최신 범위로 업데이트하고 Range를 반환.
        /// 시그니처: 배경색 == HeaderFillColor (허용 오차) AND Font.Bold == true.
        /// </summary>
        private static Excel.Range EnsureMetaRangeUpToDate(Excel.Worksheet ws, Meta meta)
        {
            int row = meta.TopRow;
            int col = meta.LeftCol;

            // 우측으로 스캔: 시그니처가 깨질 때까지
            int lastCol = col;
            int sheetLastCol = 0;
            try { sheetLastCol = ws.UsedRange?.Columns?.Count ?? 0; } catch { sheetLastCol = 0; }
            if (sheetLastCol <= 0) { try { sheetLastCol = ws.Columns.Count; } catch { sheetLastCol = 16384; } }

            // 최대 탐색: 사용열 + 여유 32
            int maxC = Math.Min(col + (sheetLastCol + 32), 16384);

            for (int c = col; c <= maxC; c++)
            {
                var cell = ws.Cells[row, c] as Excel.Range;
                if (cell == null) break;

                bool bold = false;
                int? fill = null;

                try { bold = Convert.ToBoolean(cell.Font.Bold); } catch { }
                try { fill = Convert.ToInt32(cell.Interior.Color); } catch { }

                // 시그니처 만족?
                bool isHeaderCell = bold && (fill.HasValue && ColorsClose(fill.Value, HeaderFillColor));

                if (isHeaderCell) lastCol = c;
                else break; // 연속 구간 종료
            }

            int newColCount = Math.Max(1, lastCol - col + 1);

            // 기존 ColCount와 다르면 RefersTo 갱신
            if (newColCount != meta.ColCount)
            {
                var newHeader = ws.Range[ws.Cells[row, col], ws.Cells[row, col + newColCount - 1]];
                string refersTo = "=" + newHeader.get_Address(true, true, Excel.XlReferenceStyle.xlA1, true);

                // 시트/통합문서 스코프 모두 갱신 시도
                foreach (var n in FindAllMetaNames(ws))
                {
                    try { n.RefersTo = refersTo; } catch { }
                }

                // 최신 Range 반환
                return newHeader;
            }

            // 변경 없으면 현재 Range 반환
            return ws.Range[ws.Cells[row, col], ws.Cells[row, col + meta.ColCount - 1]];
        }

        private static bool ColorsClose(int a, int b, int tolerance = 6)
        {
            // OLE Color -> R,G,B
            Color ca = ColorTranslator.FromOle(a);
            Color cb = ColorTranslator.FromOle(b);
            return Math.Abs(ca.R - cb.R) <= tolerance
                && Math.Abs(ca.G - cb.G) <= tolerance
                && Math.Abs(ca.B - cb.B) <= tolerance;
        }

        // ---------- 스타일/테두리 ----------

        private static void ApplyHeaderBorders(Excel.Worksheet ws, Excel.Range header)
        {
            // 먼저 기존 스타일 제거
            RemoveHeaderBorders(ws, header);

            // 배경
            try { header.Interior.Color = HeaderFillColor; } catch { }

            // 텍스트: Bold + 약간 크게(+1pt), 중앙 정렬
            try
            {
                header.Font.Bold = true;
                try
                {
                    double sz = Convert.ToDouble(header.Font.Size);
                    header.Font.Size = Math.Max(8, sz + 1);
                }
                catch { }
                header.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                header.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                header.WrapText = false;
                // 너무 낮으면 살짝만 높여 시각 균형
                try { if (Convert.ToSingle(header.RowHeight) < 18f) header.RowHeight = 18f; } catch { }
            }
            catch { }

            // 상단 얇은 선
            try
            {
                var bTop = header.Borders[Excel.XlBordersIndex.xlEdgeTop];
                bTop.LineStyle = ThinLineStyle;
                bTop.Color = ColorThin;
                bTop.Weight = Excel.XlBorderWeight.xlThin;
            }
            catch { }

            // 좌측 얇은 선(첫 셀)
            try
            {
                var firstCell = header.Cells[1, 1] as Excel.Range ?? header;
                var bLeft = firstCell.Borders[Excel.XlBordersIndex.xlEdgeLeft];
                bLeft.LineStyle = ThinLineStyle;
                bLeft.Color = ColorThin;
                bLeft.Weight = Excel.XlBorderWeight.xlThin;
            }
            catch { }

            // 컬럼 구분: InsideVertical
            try
            {
                var insideV = header.Borders[Excel.XlBordersIndex.xlInsideVertical];
                insideV.LineStyle = ThinLineStyle;
                insideV.Color = ColorThin;
                insideV.Weight = Excel.XlBorderWeight.xlHairline; // 더 섬세하게
            }
            catch
            {
                // fallback: 셀별 오른쪽
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
                        bRight.Weight = Excel.XlBorderWeight.xlHairline;
                    }
                    catch { }
                }
            }

            // 하단 Medium (시각적 분리)
            try
            {
                var bBottom = header.Borders[Excel.XlBordersIndex.xlEdgeBottom];
                bBottom.LineStyle = ThickLineStyle;
                bBottom.Color = ColorBottom;
                bBottom.Weight = Excel.XlBorderWeight.xlMedium;
            }
            catch { }
        }

        private static void RemoveHeaderBorders(Excel.Worksheet ws, Excel.Range header)
        {
            try
            {
                // 상단/하단/좌/InsideVertical 제거
                try { header.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlLineStyleNone; } catch { }
                try { header.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlLineStyleNone; } catch { }

                try
                {
                    var firstCell = header.Cells[1, 1] as Excel.Range;
                    if (firstCell != null)
                        firstCell.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlLineStyleNone;
                }
                catch { }

                try { header.Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlLineStyleNone; } catch { }

                // 배경/서식 복원
                try
                {
                    header.Interior.Pattern = Excel.XlPattern.xlPatternNone;
                    try { header.Interior.ColorIndex = Excel.XlColorIndex.xlColorIndexNone; } catch { }
                }
                catch { }
                try
                {
                    header.Font.Bold = false;
                    header.HorizontalAlignment = Excel.XlHAlign.xlHAlignGeneral;
                    header.VerticalAlignment = Excel.XlVAlign.xlVAlignBottom;
                }
                catch { }
            }
            catch { }
        }

        // ---------- 청소 유틸 ----------

        private static void ClearRowBottomBorders(Excel.Worksheet ws, int row)
        {
            try
            {
                int lastCol = 0;
                try { lastCol = Math.Max(ws.UsedRange?.Columns?.Count ?? 0, 1); } catch { lastCol = 0; }
                if (lastCol <= 1) { try { lastCol = ws.Columns.Count; } catch { lastCol = 256; } }
                int maxColToClear = Math.Min(lastCol + 8, 16384);

                for (int c = 1; c <= maxColToClear; c++)
                {
                    try
                    {
                        var cell = ws.Cells[row, c] as Excel.Range;
                        if (cell == null) continue;
                        cell.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlLineStyleNone;
                    }
                    catch { }
                }
                try
                {
                    var rowRange = ws.Rows[row] as Excel.Range;
                    if (rowRange != null)
                        rowRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlLineStyleNone;
                }
                catch { }
            }
            catch { }
        }

        /// <summary>과거 도형 라인이 남아있으면 제거(이름 접두사 + 헤더 근처의 긴 가로 라인)</summary>
        private static void RemoveLeftoverLineShapes(Excel.Worksheet ws, Excel.Range header)
        {
            try
            {
                var toDelete = new List<string>();
                double headerTop = 0, headerW = 0;
                try { headerTop = Convert.ToDouble(header.Top); } catch { }
                try { headerW = Convert.ToDouble(header.Width); } catch { }

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
                        if (s.Type != Microsoft.Office.Core.MsoShapeType.msoLine) continue;

                        double top = 0, w = 0;
                        try { top = Convert.ToDouble(s.Top); } catch { }
                        try { w = Convert.ToDouble(s.Width); } catch { }

                        bool nearHeader = Math.Abs(top - headerTop) <= 8.0;
                        bool veryLong = (headerW > 0) ? w > Math.Max(300.0, headerW * 1.2) : w > 1500.0;
                        if (nearHeader && veryLong)
                        {
                            if (!string.IsNullOrEmpty(nm)) toDelete.Add(nm);
                            else { try { s.Name = $"DEL_TMP_{Guid.NewGuid():N}"; toDelete.Add(s.Name); } catch { } }
                        }
                    }
                    catch { }
                }

                for (int i = toDelete.Count - 1; i >= 0; i--)
                {
                    try { ws.Shapes.Item(toDelete[i]).Delete(); } catch { }
                }
            }
            catch { }
        }
    }
}
