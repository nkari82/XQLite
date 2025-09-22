// ColumnTypeRegistry.cs
using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;

namespace XQLite.AddIn
{
    internal static class XqlColumnTypeRegistry
    {
        // 이름 규칙: _XQL_COL_{SheetName}_R{row}C{col} → ="INT"/"REAL"/"TEXT"/"BOOL"/"DATE"
        private static string NameFor(Excel.Worksheet ws, int row, int col) => $"_XQL_COL_{ws.Name}_R{row}C{col}";

        public static void SetColumnType(Excel.Worksheet ws, Excel.Range headerCell, string type)
        {
            if (ws == null || headerCell == null) throw new ArgumentNullException();
            type = (type ?? "").Trim().ToUpperInvariant();
            if (!IsSupported(type)) throw new ArgumentException($"Unsupported type: {type}");
            string name = NameFor(ws, headerCell.Row, headerCell.Column);
            string refersTo = "=\"" + type + "\"";

            var existing = FindName(ws, name);
            if (existing != null) { try { existing.RefersTo = refersTo; } catch { } }
            else { try { ws.Names.Add(name, refersTo); } catch { } }

            ApplyValidationForColumn(ws, headerCell, type);

            try
            {
                var hdr = XqlSheetMetaRegistry.GetHeaderRange(ws);
                if (hdr != null) ApplyHeaderTooltips(ws, hdr);
            }
            catch { }
        }

        public static string? GetColumnType(Excel.Worksheet ws, Excel.Range headerCell)
        {
            if (ws == null || headerCell == null) return null;
            string name = NameFor(ws, headerCell.Row, headerCell.Column);
            var n = FindName(ws, name);
            if (n == null) return null;
            try
            {
                var s = n.RefersTo as string;
                if (string.IsNullOrEmpty(s)) return null;
                var m = Regex.Match(s!, @"=\s*""([^""]+)""");
                return m.Success ? m.Groups[1].Value : null;
            }
            catch { return null; }
        }

        public static void ApplyAllForHeader(Excel.Worksheet ws)
        {
            var meta = XqlSheetMetaRegistry.Get(ws);
            if (meta == null) return;

            var header = ws.Range[ws.Cells[meta.TopRow, meta.LeftCol], ws.Cells[meta.TopRow, meta.LeftCol + Math.Max(1, meta.ColCount) - 1]];
            int cols = 1;
            try { cols = Math.Max(1, Convert.ToInt32(header.Columns.Count)); } catch { cols = 1; }

            for (int i = 1; i <= cols; i++)
            {
                var cell = header.Cells[1, i] as Excel.Range;
                if (cell == null) continue;
                var ty = GetColumnType(ws, cell);
                if (string.IsNullOrEmpty(ty)) continue;
                ApplyValidationForColumn(ws, cell, ty!);
            }
        }

        public static void ClearColumnType(Excel.Worksheet ws, Excel.Range headerCell)
        {
            if (ws == null || headerCell == null) return;
            string name = NameFor(ws, headerCell.Row, headerCell.Column);
            var n = FindName(ws, name);
            if (n != null) { try { n.Delete(); } catch { } }
            ClearValidationForColumn(ws, headerCell);
        }

        /// <summary>현재 메타 헤더 범위에 대해, 저장된 메타 타입(Name 기반)으로 툴팁을 재적용.</summary>
        public static void RefreshHeaderTooltips(Excel.Worksheet ws)
        {
            var hdr = XQLite.AddIn.XqlSheetMetaRegistry.GetHeaderRange(ws);
            if (hdr == null) return;
            ApplyHeaderTooltips(ws, hdr);
        }

        // ---- internals ----
        private static bool IsSupported(string t)
            => t == "INT" || t == "REAL" || t == "TEXT" || t == "BOOL" || t == "DATE" || t == "JSON";

        private static Excel.Name? FindName(Excel.Worksheet ws, string name)
        {
            try { foreach (Excel.Name n in ws.Names) if ((n.NameLocal ?? n.Name) == name) return n; } catch { }
            try { foreach (Excel.Name n in ws.Application.ActiveWorkbook.Names) { var nm = n.NameLocal ?? n.Name; if (nm == name || nm.EndsWith("!" + name)) return n; } } catch { }
            return null;
        }

        private static void ApplyValidationForColumn(Excel.Worksheet ws, Excel.Range headerCell, string type)
        {
            int col = headerCell.Column;
            int startRow = headerCell.Row + 1;
            int lastRow; try { lastRow = Math.Min(1048576, (ws.UsedRange?.Rows?.Count ?? 20000) + 5000); } catch { lastRow = 50000; }

            var range = ws.Range[ws.Cells[startRow, col], ws.Cells[lastRow, col]];
            range.Validation.InputTitle = $"[{type}] column";
            range.Validation.ShowInput = true;
            try { range.Validation.Delete(); } catch { }

            try
            {
                // A1 주소(상대참조) 생성: 검증 수식은 A1 기준으로 작성해야 셀마다 자동 이동됨
                string a1 = (ws.Cells[startRow, col] as Excel.Range)!.Address[RowAbsolute: false, ColumnAbsolute: false, ReferenceStyle: Excel.XlReferenceStyle.xlA1];

                switch (type)
                {
                    case "INT":
                        range.NumberFormat = "0";
                        range.Validation.Add(
                            Excel.XlDVType.xlValidateWholeNumber,
                            Excel.XlDVAlertStyle.xlValidAlertStop,
                            Excel.XlFormatConditionOperator.xlBetween,
                            -2147483648, 2147483647);
                        range.Validation.IgnoreBlank = true;
                        range.Validation.ErrorMessage = "정수만 입력하세요.";
                        range.Validation.InputMessage = "정수만 허용 (예: 0, -42, 123456)";
                        break;

                    case "REAL":
                        // ✅ locale/지수표기 문제 회피: 'Custom'으로 ISNUMBER로만 검증
                        range.NumberFormat = "General"; // 지수표기도 자연스럽게 표시
                        string fReal = $"=OR(ISBLANK({a1}),ISNUMBER({a1}))";
                        range.Validation.Add(
                            Excel.XlDVType.xlValidateCustom,
                            Excel.XlDVAlertStyle.xlValidAlertStop,
                            Type.Missing, fReal, Type.Missing);
                        range.Validation.IgnoreBlank = true;
                        range.Validation.ErrorMessage = "실수(숫자)만 입력하세요. (예: 3.14, 1E-6)";
                        range.Validation.InputMessage = "실수(숫자)만 허용 (예: 3.14, -0.001, 1E-6)";
                        break;

                    case "TEXT":
                        range.NumberFormat = "@";
                        range.Validation.InputMessage = "문자열 입력";
                        // 텍스트는 별도 검증 없이 허용
                        break;

                    case "BOOL":
                        range.NumberFormat = "General";
                        range.Validation.Add(
                            Excel.XlDVType.xlValidateList,
                            Excel.XlDVAlertStyle.xlValidAlertStop,
                            Excel.XlFormatConditionOperator.xlBetween,
                            "\"0,1,TRUE,FALSE,Yes,No,True,False\"",
                            Type.Missing);
                        range.Validation.IgnoreBlank = true;
                        range.Validation.InCellDropdown = true;
                        range.Validation.ErrorMessage = "0/1 또는 TRUE/FALSE만 허용됩니다.";
                        range.Validation.InputMessage = "TRUE/FALSE 또는 0/1";
                        break;

                    case "DATE":
                        range.NumberFormat = "yyyy-mm-dd";
                        range.Validation.Add(
                            Excel.XlDVType.xlValidateDate,
                            Excel.XlDVAlertStyle.xlValidAlertStop,
                            Excel.XlFormatConditionOperator.xlBetween,
                            new DateTime(1900, 1, 1), new DateTime(9999, 12, 31));
                        range.Validation.IgnoreBlank = true;
                        range.Validation.ErrorMessage = "유효한 날짜를 입력하세요.";
                        range.Validation.InputMessage = "날짜 (yyyy-mm-dd)";
                        break;

                    case "JSON":
                        // ✅ JSON(간단 검증): 비어있거나, {..} 또는 [..] 로 감싸진 텍스트
                        range.NumberFormat = "@";
                        try { range.Font.Name = "Consolas"; } catch { } // 가독성(선택)
                        string fJson =
                            $"=OR(LEN(TRIM({a1}))=0," +
                            $"AND(LEFT(TRIM({a1}),1)=\"{{\",RIGHT(TRIM({a1}),1)=\"}}\")," +
                            $"AND(LEFT(TRIM({a1}),1)=\"[\",RIGHT(TRIM({a1}),1)=\"]\"))";
                        range.Validation.Add(
                            Excel.XlDVType.xlValidateCustom,
                            Excel.XlDVAlertStyle.xlValidAlertStop,
                            Type.Missing, fJson, Type.Missing);
                        range.Validation.IgnoreBlank = true;
                        range.Validation.ErrorMessage = "유효한 JSON이어야 합니다. (예: {\"k\":1} 또는 [1,2])";
                        range.Validation.InputMessage = "JSON 텍스트 (예: {\"k\":1} 또는 [1,2])";
                        break;
                }
            }
            catch { /* best-effort */ }
        }

        private static void ClearValidationForColumn(Excel.Worksheet ws, Excel.Range headerCell)
        {
            int col = headerCell.Column;
            int startRow = headerCell.Row + 1;
            int lastRow; try { lastRow = Math.Min(1048576, (ws.UsedRange?.Rows?.Count ?? 20000) + 5000); } catch { lastRow = 50000; }
            var range = ws.Range[ws.Cells[startRow, col], ws.Cells[lastRow, col]];
            try { range.Validation.Delete(); } catch { }
            try { range.NumberFormat = "General"; } catch { }
        }

        // ColumnTypeRegistry 내부

        // ColumnTypeRegistry 내부의 툴팁 유틸을 다음으로 교체

        private static void ApplyHeaderTooltips(Excel.Worksheet ws, Excel.Range header)
        {
            int colCount = 1;
            try { colCount = Math.Max(1, Convert.ToInt32(header.Columns.Count)); } catch { colCount = 1; }

            for (int i = 1; i <= colCount; i++)
            {
                var cell = header.Cells[1, i] as Excel.Range;
                if (cell == null) continue;

                string colName = Convert.ToString(cell.Value2) ?? "";
                if (string.IsNullOrWhiteSpace(colName)) { SafeClearNote(cell); continue; }

                var ty = GetColumnType(ws, cell) ?? "TEXT";

                string hint = ty.ToUpperInvariant() switch
                {
                    "INT" => "Type: INT\n예) 0, -42, 123456",
                    "REAL" => "Type: REAL\n예) 3.14, -0.001, 1E-6",
                    "BOOL" => "Type: BOOL\n예) TRUE/FALSE, 0/1",
                    "DATE" => "Type: DATE (yyyy-mm-dd)\n예) 2025-09-22",
                    "JSON" => "Type: JSON\n예) {\"k\":1} 또는 [1,2]",
                    _ => "Type: TEXT"
                };

                SetNoteLegacy(cell, hint);
            }
        }

        private static void SetNoteLegacy(Excel.Range cell, string text)
        {
            SafeClearNote(cell);
            try
            {
                // 구(舊) 코멘트 API
                cell.AddComment(text);
                try
                {
                    // 자동 줄바꿈/가독성 소폭 개선
                    cell.Comment.Shape.TextFrame.Characters().Font.Size = 9;
                    cell.Comment.Shape.TextFrame.AutoSize = true;
                    cell.Comment.Visible = false; // 마우스오버/선택 시만 보이게
                }
                catch { }
            }
            catch
            {
                // 조직 정책 등으로 메모가 금지된 환경이면 무시
            }
        }

        /// <summary>메모(Comment)만 안전하게 제거</summary>
        private static void SafeClearNote(Excel.Range cell)
        {
            try
            {
                // 일부 버전은 ClearComments 지원
                cell.ClearComments();
                return;
            }
            catch { /* fallthrough */ }

            try
            {
                if (cell.Comment != null) cell.Comment.Delete();
            }
            catch { /* ignore */ }
        }

    }
}
