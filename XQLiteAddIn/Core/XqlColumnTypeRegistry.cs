// ColumnTypeRegistry.cs
using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;

#if false
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

        private static string Trunc(string s, int max)
        {
            if (string.IsNullOrEmpty(s)) return s;
            return s.Length <= max ? s : s.Substring(0, max);
        }

        private static void ApplyValidationForColumn(Excel.Worksheet ws, Excel.Range headerCell, string type)
        {
            int col = headerCell.Column;
            int startRow = headerCell.Row + 1;

            int lastRow;
            try { lastRow = Math.Min(1048576, (ws.UsedRange?.Rows?.Count ?? 20000) + 5000); }
            catch { lastRow = 50000; }

            var range = ws.Range[ws.Cells[startRow, col], ws.Cells[lastRow, col]];
            try { range.Validation.Delete(); } catch { /* ignore */ }

            // A1 상대주소와 지역 구분자 확보
            string a1 = (ws.Cells[startRow, col] as Excel.Range)!
                .Address[RowAbsolute: false, ColumnAbsolute: false, ReferenceStyle: Excel.XlReferenceStyle.xlA1];
            string sep = ws.Application.International[Excel.XlApplicationInternational.xlListSeparator] as string ?? ",";

            // 헬퍼: Custom 규칙 추가(포뮬러1 only) — 로캘 구분자 재시도
            void AddCustom(string formulaBody)
            {
                string f = "=" + formulaBody.Replace(",", sep);
                try
                {
                    range.Validation.Add(
                        Excel.XlDVType.xlValidateCustom,
                        Excel.XlDVAlertStyle.xlValidAlertStop,
                        Excel.XlFormatConditionOperator.xlBetween,
                        f, Type.Missing); // <-- formula1 여기에!
                }
                catch
                {
                    // 혹시 반대로 설정된 환경 대비(거의 필요 없음)
                    string alt = "=" + formulaBody.Replace(";", ",");
                    range.Validation.Add(
                        Excel.XlDVType.xlValidateCustom,
                        Excel.XlDVAlertStyle.xlValidAlertStop,
                        Excel.XlFormatConditionOperator.xlBetween,
                        alt, Type.Missing);
                }
            }

            switch (type.ToUpperInvariant())
            {
                case "INT":
                    range.NumberFormat = "0";
                    range.Validation.Add(
                        Excel.XlDVType.xlValidateWholeNumber,
                        Excel.XlDVAlertStyle.xlValidAlertStop,
                        Excel.XlFormatConditionOperator.xlBetween,
                        int.MinValue, int.MaxValue);
                    break;

                case "REAL":
                    range.NumberFormat = "General";
                    // ISNUMBER(A1) 또는 빈칸 허용
                    AddCustom($"OR(ISBLANK({a1}),ISNUMBER({a1}))");
                    break;

                case "TEXT":
                    range.NumberFormat = "@";
                    // 항상 TRUE인 커스텀(메시지만 쓰기 위함)
                    AddCustom("TRUE");
                    break;

                case "BOOL":
                    range.NumberFormat = "General";
                    // 리스트는 로캘 구분자 영향 안 받는 경우가 많지만 재시도 가드
                    try
                    {
                        range.Validation.Add(
                            Excel.XlDVType.xlValidateList,
                            Excel.XlDVAlertStyle.xlValidAlertStop,
                            Excel.XlFormatConditionOperator.xlBetween,
                            "\"0,1,TRUE,FALSE,Yes,No,True,False\"",
                            Type.Missing);
                    }
                    catch
                    {
                        range.Validation.Add(
                            Excel.XlDVType.xlValidateList,
                            Excel.XlDVAlertStyle.xlValidAlertStop,
                            Excel.XlFormatConditionOperator.xlBetween,
                            "\"0;1;TRUE;FALSE;Yes;No;True;False\"",
                            Type.Missing);
                    }
                    range.Validation.InCellDropdown = true;
                    break;

                case "DATE":
                    range.NumberFormat = "yyyy-mm-dd";
                    range.Validation.Add(
                        Excel.XlDVType.xlValidateDate,
                        Excel.XlDVAlertStyle.xlValidAlertStop,
                        Excel.XlFormatConditionOperator.xlBetween,
                        new DateTime(1900, 1, 1), new DateTime(9999, 12, 31));
                    break;

                case "JSON":
                    range.NumberFormat = "@";
                    try { range.Font.Name = "Consolas"; } catch { }
                    // 빈칸 or {..} or [..]
                    AddCustom(
                        $"OR(LEN(TRIM({a1}))=0," +
                        $"AND(LEFT(TRIM({a1}),1)=\"{{\",RIGHT(TRIM({a1}),1)=\"}}\")," +
                        $"AND(LEFT(TRIM({a1}),1)=\"[\",RIGHT(TRIM({a1}),1)=\"]\"))");
                    break;

                default:
                    AddCustom("TRUE");
                    break;
            }

            // (중요) 규칙 만든 뒤에 메시지/옵션
            try
            {
                string title = Trunc($"[{type}] column", 32);

                string msg = type.ToUpperInvariant() switch
                {
                    "INT" => "정수만 허용 (예: 0, -42, 123456)",
                    "REAL" => "실수(숫자)만 허용 (예: 3.14, -0.001, 1E-6)",
                    "TEXT" => "문자열 입력",
                    "BOOL" => "TRUE/FALSE 또는 0/1",
                    "DATE" => "날짜 (yyyy-mm-dd)",
                    "JSON" => "JSON 텍스트 (예: {\"k\":1} 또는 [1,2])",
                    _ => "입력 도움말"
                };
                if (msg.Length > 255) msg = Trunc(msg, 255);

                range.Validation.InputTitle = title;
                range.Validation.InputMessage = msg;
                range.Validation.ShowInput = true;
                range.Validation.IgnoreBlank = true;
            }
            catch { /* 일부 환경 제한 → 무시 */ }
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
#endif