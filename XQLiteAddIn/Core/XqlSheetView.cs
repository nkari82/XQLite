// XqlSheetView.cs
using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using MessageBox = System.Windows.Forms.MessageBox;

namespace XQLite.AddIn
{
    /// <summary>
    /// 시트 UI 전담: 메타 헤더(시트당 하나) 삽입/삭제/새로고침/정보 + 주석/테두리/데이터 검증.
    /// XqlSheet 인스턴스를 통해 메타 레지스트리 및 유틸 사용.
    /// </summary>
    internal static class XqlSheetView
    {
        private const string Caption = "XQLite";

        // ───────────────────────── Public API (Ribbon에서 호출)
        // 리본에서 호출하는 최종 진입점
        public static bool InstallHeader()
        {
            var app = (Excel.Application)ExcelDnaUtil.Application;
            Excel.Worksheet? ws = null;
            Excel.Range? candidate = null;   // ← 바깥에 선언(마지막에 해제)
            try
            {
                ws = (Excel.Worksheet)app.ActiveSheet;
                if (ws == null) return false;

                var sheet = XqlAddIn.Sheet;
                if (sheet == null) { MessageBox.Show("Sheet service not ready.", "XQLite"); return false; }

                candidate = GetHeaderOrFallback(ws);
                if (candidate == null)
                {
                    MessageBox.Show("헤더 후보를 찾을 수 없습니다.", "XQLite",
                                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return false;
                }

                // 1) 기존 헤더 마커 확인 → 다른 위치면 중복 설치 경고
                if (XqlSheet.TryGetHeaderMarker(ws, out var old) && !XqlSheet.IsSameRange(old, candidate))
                {
                    XqlCommon.ReleaseCom(old);
                    MessageBox.Show("이 시트에는 이미 헤더가 설치되어 있습니다.\r\n(다른 위치에 중복 설치할 수 없습니다)",
                                    "XQLite", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return false;
                }
                XqlCommon.ReleaseCom(old);

                // 2) 메타 동기화
                var names = BuildHeaderNames(candidate);
                var sm = sheet.GetOrCreateSheet(ws.Name);
                sheet.EnsureColumns(ws.Name, names);

                // 3) 툴팁/테두리/검증
                var tips = BuildHeaderTooltips(sm, candidate);
                SetHeaderTooltips(candidate, tips);
                ApplyHeaderOutlineBorder(candidate);
                ApplyDataValidationForHeader(ws, candidate, sm);

                // 4) 마커 (단 하나)
                XqlSheet.SetHeaderMarker(ws, candidate);
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("InstallHeader failed: " + ex.Message, "XQLite");
                return false;
            }
            finally
            {
                XqlCommon.ReleaseCom(candidate);
                XqlCommon.ReleaseCom(ws);
            }
        }

        // 컬럼 타입 → 툴팁 문자열
        private static string ColumnTooltipFor(ColumnType ct)
        {
            // XqlSheet.cs의 ColumnType.ToTooltip() 재사용
            return ct?.ToTooltip() ?? "TEXT • NULL OK";
        }

        // 이름 매칭이 안 될 때 인덱스 기반 폴백
        private static string ColumnTooltipFallback(SheetMeta sm, int keyIndex)
        {
            if (sm == null || sm.Columns == null || sm.Columns.Count == 0) return "TEXT • NULL OK";
            // Dictionary라도 .NET 최신 런타임에서 삽입순서를 유지하므로 안전한 폴백으로 사용
            if (keyIndex >= 0 && keyIndex < sm.Columns.Count)
            {
                var ct = sm.Columns.ElementAt(keyIndex).Value;
                return ColumnTooltipFor(ct);
            }
            return "TEXT • NULL OK";
        }

        // 헤더 Range로부터 “이름 우선, 위치 폴백” 툴팁 맵 구성
        internal static IReadOnlyDictionary<string, string> BuildHeaderTooltips(SheetMeta sm, Excel.Range header)
        {
            var dict = new Dictionary<string, string>(StringComparer.Ordinal);

            // 1) 이름 기반
            foreach (var kv in sm.Columns)
                dict[kv.Key] = ColumnTooltipFor(kv.Value);

            // 2) 위치 기반 폴백(@1, @2, …): 이름이 없는 칸/미등록 이름 대응
            int idx = 0;
            foreach (Excel.Range cell in header.Cells)
            {
                try
                {
                    var name = (cell.Value2 as string)?.Trim();
                    if (!string.IsNullOrEmpty(name) && dict.ContainsKey(name!))
                    {
                        idx++;
                        continue; // 이름 매칭이 되면 위치 폴백 불필요
                    }
                    dict[$"@{++idx}"] = ColumnTooltipFallback(sm, idx - 1);
                }
                finally { XqlCommon.ReleaseCom(cell); }
            }
            return dict;
        }

        public static void RefreshHeader()
        {
            var app = (Excel.Application)ExcelDnaUtil.Application;
            Excel.Worksheet? ws = null; Excel.Range? header = null;
            try
            {
                ws = (Excel.Worksheet)app.ActiveSheet;
                if (ws == null) return;

                header = GetHeaderOrFallback(ws);
                if (header == null) { MessageBox.Show("헤더를 찾을 수 없습니다.", Caption); return; }

                // 헤더가 이동되었으면 마커를 새 위치로 재바인딩
                RebindMarkerIfMoved(ws, header);

                var sheet = XqlAddIn.Sheet!;
                var sm = sheet.GetOrCreateSheet(ws.Name);

                var tips = BuildHeaderTooltips(sm, header);
                SetHeaderTooltips(header, tips);
                ApplyHeaderOutlineBorder(header);
                ApplyDataValidationForHeader(ws, header, sm);
            }
            catch (Exception ex)
            {
                MessageBox.Show("RefreshMetaHeader failed: " + ex.Message, Caption);
            }
            finally { XqlCommon.ReleaseCom(header); XqlCommon.ReleaseCom(ws); }
        }

        public static void RemoveHeader()
        {
            var app = (Excel.Application)ExcelDnaUtil.Application;
            Excel.Worksheet? ws = null; Excel.Range? hdr = null; Excel.Range? sel = null;
            try
            {
                ws = (Excel.Worksheet)app.ActiveSheet;
                if (ws == null) return;
                if (!XqlSheet.TryGetHeaderMarker(ws, out hdr))
                {
                    sel = GetSelection(ws);
                    hdr = ResolveHeader(ws, sel, XqlAddIn.Sheet!) ?? XqlSheet.GetHeaderRange(ws);
                }
                ClearHeaderUi(ws, hdr, removeMarker: true);
            }
            catch (Exception ex)
            {
                MessageBox.Show("RemoveMetaHeader failed: " + ex.Message, Caption);
            }
            finally { XqlCommon.ReleaseCom(sel); XqlCommon.ReleaseCom(hdr); XqlCommon.ReleaseCom(ws); }
        }

        public static void ShowHeaderInfo()
        {
            var app = (Excel.Application)ExcelDnaUtil.Application;
            Excel.Worksheet? ws = null;

            Excel.Range? header = null;
            try
            {
                ws = (Excel.Worksheet)app.ActiveSheet;
                if (ws == null) return;


                header = GetHeaderOrFallback(ws);
                if (header == null) { MessageBox.Show("헤더를 찾을 수 없습니다.", Caption); return; }

                // 선택이 실제 헤더 행이라면 마커를 그 위치로 재바인딩 (헤더 이동 반영)
                RebindMarkerIfMoved(ws, header);

                var sheet = XqlAddIn.Sheet!;
                var sm = sheet.GetOrCreateSheet(ws.Name);

                // 헤더 정보 구성 (전 컬럼 순회)
                var sb = new System.Text.StringBuilder();
                var addr = header.Address[true, true, Excel.XlReferenceStyle.xlA1, false];
                var startColLetter = XqlCommon.ColumnIndexToLetter(header.Column);
                var startRow = header.Row;
                sb.AppendLine($"{ws.Name}!{addr}");
                sb.AppendLine($"Start: Col {startColLetter} ({header.Column}), Row {startRow}  |  Data starts @ {startRow + 1}");
                sb.AppendLine("");

                for (int i = 1; i <= header.Columns.Count; i++)
                {
                    Excel.Range? h = null;
                    try
                    {
                        h = (Excel.Range)header.Cells[1, i];
                        var colName = (h.Value2 as string)?.Trim();
                        if (string.IsNullOrEmpty(colName))
                            colName = XqlCommon.ColumnIndexToLetter(h.Column);

                        if (sm.Columns.TryGetValue(colName!, out var ct))
                            sb.AppendLine($"{ws.Name}.{colName}\r\n{ct.ToTooltip()}");
                        else
                            sb.AppendLine($"{ws.Name}.{colName}\r\nTEXT • NULL OK");
                        sb.AppendLine("");
                    }
                    finally { XqlCommon.ReleaseCom(h); }
                }

                MessageBox.Show(sb.ToString().TrimEnd(), Caption);
            }
            catch (Exception ex)
            {
                MessageBox.Show("ShowMetaHeaderInfo failed: " + ex.Message, Caption);
            }
            finally { XqlCommon.ReleaseCom(header); XqlCommon.ReleaseCom(ws); }
        }

        // ───────────────────────── Header Resolve / Marker

        public static Excel.Range? ResolveHeader(Excel.Worksheet ws, Excel.Range? sel, XqlSheet sheet)
        {
            if (ws == null) return null;
            if (sel != null)
            {
                var loSel = sel.ListObject ?? XqlSheet.FindListObjectContaining(ws, sel);
                if (loSel?.HeaderRowRange != null) return loSel.HeaderRowRange;

                int r = sel.Row; int c1 = sel.Column; int c2 = c1 + sel.Columns.Count - 1;
                return ws.Range[ws.Cells[r, c1], ws.Cells[r, c2]];
            }
            return null;
        }

        private static Excel.Range? GetSelection(Excel.Worksheet ws)
        {
            try { return (Excel.Range)ws.Application.Selection; }
            catch { return null; }
        }

        // ───────────────────────── UI Helpers (Comments / Borders / Tooltips / Validation)

        private static List<string> BuildHeaderNames(Excel.Range header)
        {
            var names = new List<string>(header.Columns.Count);
            var v = header.Value2 as object[,];
            if (v != null)
            {
                int cols = header.Columns.Count;
                for (int c = 1; c <= cols; c++)
                {
                    string name = (Convert.ToString(v[1, c]) ?? string.Empty).Trim();
                    if (string.IsNullOrEmpty(name))
                    {
                        var cell = (Excel.Range)header.Cells[1, c];
                        name = XqlCommon.ColumnIndexToLetter(cell.Column);
                        XqlCommon.ReleaseCom(cell);
                    }
                    names.Add(name);
                }
                return names;
            }

            foreach (Excel.Range cell in header.Cells)
            {
                string name = (Convert.ToString(cell.Value2) ?? string.Empty).Trim();
                if (string.IsNullOrEmpty(name))
                    name = XqlCommon.ColumnIndexToLetter(cell.Column);
                names.Add(name);
                XqlCommon.ReleaseCom(cell);
            }
            return names;
        }

        internal static void SetHeaderTooltips(Excel.Range header, IReadOnlyDictionary<string, string> tips)
        {
            int pos = 0;
            foreach (Excel.Range cell in header.Cells)
            {
                try
                {
                    var nameKey = (cell.Value2 as string)?.Trim();
                    var posKey = $"@{++pos}";
                    if (!tips.TryGetValue(nameKey ?? string.Empty, out var tip))
                        tips.TryGetValue(posKey, out tip);
                    if (string.IsNullOrEmpty(tip)) continue;

                    // 변경시에만 갱신 + 512자 절삭 + 숨김
                    Excel.Comment? c = null;
                    try
                    {
                        c = cell.Comment;
                        var old = c is null ? null : c.Text() as string;
                        var safe = tip.Length <= 512 ? tip : (tip.Substring(0, 509) + "...");
                        if (!string.Equals(old?.Trim(), safe, StringComparison.Ordinal))
                        {
                            try { c?.Delete(); } catch { }
                            c = cell.AddComment(safe);
                            try { if (c != null) c.Visible = false; } catch { }
                        }
                    }
                    finally { XqlCommon.ReleaseCom(c); }
                }
                finally { XqlCommon.ReleaseCom(cell); }
            }
        }

        // 헤더 1행이 편집되면 툴팁 재적용
        public static void RefreshTooltipsIfHeaderEdited(Excel.Worksheet ws, Excel.Range target)
        {
            if (ws == null || target == null) return;
            var sheet = XqlAddIn.Sheet;
            if (sheet == null) return;

            Excel.Range? header;
            if (!XqlSheet.TryGetHeaderMarker(ws, out header)) return;

            try
            {
                var app = ws.Application;
                var inter = app.Intersect(header, target);
                if (inter == null) return;

                if (sheet.TryGetSheet(ws.Name, out var sm))
                {
                    var tips = BuildHeaderTooltips(sm, header);
                    SetHeaderTooltips(header, tips);
                }
                XqlCommon.ReleaseCom(inter);
            }
            finally { XqlCommon.ReleaseCom(header); }
        }

        internal static void ApplyDataValidationForHeader(Excel.Worksheet ws, Excel.Range header, SheetMeta sm)
        {
            var lo = XqlSheet.FindListObjectContaining(ws, header);
            if (lo?.HeaderRowRange != null)
            {
                for (int i = 1; i <= header.Columns.Count; i++)
                {
                    Excel.Range? h = null; Excel.Range? body = null; Excel.Range? rng = null;
                    try
                    {
                        h = (Excel.Range)header.Cells[1, i];
                        string? name = (h.Value2 as string)?.Trim();
                        if (string.IsNullOrEmpty(name)) name = XqlCommon.ColumnIndexToLetter(h.Column);

                        // 메타에 없더라도 일단 '열 전체'에 완화 규칙(Any)로 깔고, 있으면 실제 타입으로 다시 덮어쓴다.
                        rng = ColBelowToEnd(ws, h);
                        ApplyValidationForKind(rng, ColumnKind.Text /*완화; 먼저 깨끗이 덮고*/);

                        if (!sm.Columns.TryGetValue(name!, out var ct))
                            continue;

                        // 실제 타입으로 다시 덮어쓰기
                        ApplyValidationForKind(rng, ct.Kind);
                    }
                    finally { XqlCommon.ReleaseCom(h); XqlCommon.ReleaseCom(body); XqlCommon.ReleaseCom(rng); }
                }
                return;
            }

            // ── 표 바깥(일반 범위) 폴백 ──
            for (int i = 1; i <= header.Columns.Count; i++)
            {
                Excel.Range? h = null; Excel.Range? col = null;
                try
                {
                    h = (Excel.Range)header.Cells[1, i];
                    var name = (h.Value2 as string)?.Trim();
                    if (string.IsNullOrEmpty(name)) name = XqlCommon.ColumnIndexToLetter(h.Column);

                    col = ColBelowToEnd(ws, h); // ✅ UsedRange 대신 시트 끝까지
                    if (!string.IsNullOrEmpty(name) && sm.Columns.TryGetValue(name!, out var ct))
                        ApplyValidationForKind(col, ct.Kind);
                    else
                        ApplyValidationForKind(col, ColumnKind.Text /*완화*/);
                }
                finally { XqlCommon.ReleaseCom(h); XqlCommon.ReleaseCom(col); }
            }
        }

        private static void ApplyHeaderOutlineBorder(Excel.Range header)
        {
            Excel.Borders? bs = null;
            try
            {
                bs = header.Borders;
                var idxs = new[]
                {
                    Excel.XlBordersIndex.xlEdgeLeft,
                    Excel.XlBordersIndex.xlEdgeTop,
                    Excel.XlBordersIndex.xlEdgeRight,
                    Excel.XlBordersIndex.xlEdgeBottom
                };
                foreach (var idx in idxs)
                {
                    var b = bs[idx];
                    try
                    {
                        b.LineStyle = Excel.XlLineStyle.xlContinuous;
                        b.Weight = Excel.XlBorderWeight.xlThin;
                    }
                    finally
                    {
                        XqlCommon.ReleaseCom(b);
                    }
                }
            }
            catch { }
            finally
            {
                XqlCommon.ReleaseCom(bs);
            }
        }

        // ClearHeaderUi(...)
        private static void ClearHeaderUi(Excel.Worksheet ws, Excel.Range? header, bool removeMarker = false)
        {
            if (header == null) header = XqlSheet.GetHeaderRange(ws);

            // 1) 헤더 툴팁(코멘트) 제거
            foreach (Excel.Range cell in header.Cells)
            {
                try
                {
                    try { cell.ClearComments(); } catch { try { cell.Comment?.Delete(); } catch { } }
                }
                finally { XqlCommon.ReleaseCom(cell); }
            }

            // 2) 헤더 테두리/내부선 제거
            try
            {
                var bs = header.Borders;
                var idxs = new[]
                       {
           Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBordersIndex.xlEdgeTop,
           Excel.XlBordersIndex.xlEdgeRight, Excel.XlBordersIndex.xlEdgeBottom,
           Excel.XlBordersIndex.xlInsideHorizontal, Excel.XlBordersIndex.xlInsideVertical
       };
                foreach (var idx in idxs)
                {
                    var b = bs[idx]; try { b.LineStyle = Excel.XlLineStyle.xlLineStyleNone; } finally { XqlCommon.ReleaseCom(b); }
                }
                XqlCommon.ReleaseCom(bs);
            }
            catch { }

            // 3) 헤더 아래 열 전체의 데이터 유효성 제거 (완전 초기화)
            try
            {
                foreach (Excel.Range h in header.Cells)
                {
                    Excel.Range? col = null;
                    try
                    {
                        var first = (Excel.Range)h.Offset[1, 0];
                        var last = ws.Cells[ws.Rows.Count, h.Column];
                        col = ws.Range[first, last];
                        try { col.Validation.Delete(); } catch { }
                        XqlCommon.ReleaseCom(first); XqlCommon.ReleaseCom(last);
                    }
                    finally { XqlCommon.ReleaseCom(col); XqlCommon.ReleaseCom(h); }
                }
            }
            catch { }

            if (removeMarker) XqlSheet.ClearHeaderMarker(ws);
        }

        private static Excel.Range ColBelowToEnd(Excel.Worksheet ws, Excel.Range h)
        {
            var first = (Excel.Range)h.Offset[1, 0];
            var last = ws.Cells[ws.Rows.Count, h.Column];
            var rng = ws.Range[first, last];
            XqlCommon.ReleaseCom(first);
            XqlCommon.ReleaseCom(last);
            return rng;
        }

        private static void ApplyValidationForKind(Excel.Range rng, ColumnKind kind)
        {
            try
            {
                // 기존 규칙이 있으면 지우고 새로 깝니다 (중복/병합으로 인한 Add 실패 방지
                try { rng.Validation.Delete(); } catch { }

                // 기존 Validation.Add 분기들 유지
                switch (kind)
                {
                    case ColumnKind.Int:
                        rng.Validation.Add(
                            Excel.XlDVType.xlValidateWholeNumber,
                            Excel.XlDVAlertStyle.xlValidAlertStop,
                            Excel.XlFormatConditionOperator.xlBetween,
                            -2147483648, 2147483647);
                        rng.Validation.IgnoreBlank = true;
                        rng.Validation.ErrorTitle = "정수만 허용";
                        rng.Validation.ErrorMessage = "이 열은 정수만 입력할 수 있습니다.";
                        break;

                    case ColumnKind.Real:
                        rng.Validation.Add(
                            Excel.XlDVType.xlValidateDecimal,
                            Excel.XlDVAlertStyle.xlValidAlertStop,
                            Excel.XlFormatConditionOperator.xlBetween,
                            -1e307, 1e307);
                        rng.Validation.IgnoreBlank = true;
                        rng.Validation.ErrorTitle = "숫자만 허용";
                        rng.Validation.ErrorMessage = "이 열은 숫자(실수)만 입력할 수 있습니다.";
                        break;

                    case ColumnKind.Bool:
                        rng.Validation.Add(
                            Excel.XlDVType.xlValidateList,
                            Excel.XlDVAlertStyle.xlValidAlertStop,
                            Excel.XlFormatConditionOperator.xlBetween,
                            "TRUE,FALSE");
                        rng.Validation.IgnoreBlank = true;
                        rng.Validation.ErrorTitle = "TRUE/FALSE만 허용";
                        rng.Validation.ErrorMessage = "이 열은 TRUE 또는 FALSE만 입력할 수 있습니다.";
                        break;

                    case ColumnKind.Date:
                        rng.Validation.Add(
                            Excel.XlDVType.xlValidateDate,
                            Excel.XlDVAlertStyle.xlValidAlertStop,
                            Excel.XlFormatConditionOperator.xlBetween,
                            DateTime.MinValue, DateTime.MaxValue);
                        rng.Validation.IgnoreBlank = true;
                        rng.Validation.ErrorTitle = "날짜만 허용";
                        rng.Validation.ErrorMessage = "이 열은 날짜만 입력할 수 있습니다.";
                        break;

                    case ColumnKind.Json:
                    case ColumnKind.Text:
                    default:
                        // 제한 없음 (JSON은 서버/Change 이벤트에서 보조 검사)
                        break;
                }

                // ✨ 중요: 에러 박스/인풋풍선 설정
                try { rng.Validation.ShowError = true; } catch { }
                try { rng.Validation.ShowInput = false; } catch { }
            }
            catch
            {
                // 병합/빈 범위 등 실패 가능 – 침묵
            }
        }

        // XqlSheetView.cs 내부 (클래스 상단 private static 영역)
        /// <summary>
        /// 헤더 범위를 얻는다. 우선순위: 마커 → (선택 기반) ResolveHeader → Fallback(GetHeaderRange)
        /// </summary>
        private static Excel.Range? GetHeaderOrFallback(Excel.Worksheet ws)
        {
            Excel.Range? hdr;
            if (XqlSheet.TryGetHeaderMarker(ws, out hdr))
                return hdr;

            Excel.Range? sel = null; Excel.Range? guess = null;
            try
            {
                sel = GetSelection(ws);
                // ResolveHeader가 있으면 활용, 실패 시 1행 Fallback
                guess = ResolveHeader(ws, sel, XqlAddIn.Sheet!) ?? XqlSheet.GetHeaderRange(ws);
                return guess;
            }
            finally { XqlCommon.ReleaseCom(sel); /* guess는 반환 */ }
        }

        /// <summary>
        /// 선택된 범위가 '헤더처럼 보이면' 마커를 그 위치로 옮긴다(열 개수 ≥1, 전부 비어있지 않음).
        /// 기존 마커가 있고 주소가 다르면 자동 재바인딩.
        /// </summary>
        private static void RebindMarkerIfMoved(Excel.Worksheet ws, Excel.Range candidate)
        {
            Excel.Range? old = null;
            try
            {
                if (XqlSheet.TryGetHeaderMarker(ws, out old))
                {
                    if (!XqlSheet.IsSameRange(old, candidate))
                        XqlSheet.SetHeaderMarker(ws, candidate); // ← 새 위치로 재바인딩
                }
                else
                {
                    XqlSheet.SetHeaderMarker(ws, candidate);       // 마커 없으면 새로 생성
                }
            }
            finally { XqlCommon.ReleaseCom(old); }
        }


        // (Option) JSON/특수 타입 보조 검사 – 메시지 박스 직접 띄우기
        // Worksheet_Change if (XqlSheet.TryGetSheet(ws.Name, out var sm)) XqlSheetView.EnforceTypeOnChange(ws, target, sm);
#if false
        internal static void EnforceTypeOnChange(Excel.Worksheet ws, Excel.Range target, SheetMeta sm)
        {
            // 헤더/표 찾기
            var lo = XqlSheet.FindListObjectContaining(ws, target);
            if (lo?.HeaderRowRange == null || lo.DataBodyRange == null) return;

            // 대상이 바디와 교차?
            if (ws.Application.Intersect(lo.DataBodyRange, target) == null) return;

            foreach (Excel.Range cell in target.Cells)
            {
                Excel.ListColumn? lc = null; Excel.Range? h = null;
                try
                {
                    lc = lo.ListColumns[cell.Column - lo.DataBodyRange.Column + 1];
                    h = (Excel.Range?)(lc?.Range.Cells[1, 1]);
                    var name = (h?.Value2 as string)?.Trim();
                    if (string.IsNullOrEmpty(name) || !sm.Columns.TryGetValue(name!, out var ct)) continue;

                    var txt = (cell.Value2 is string s) ? s : cell.Value2?.ToString();
                    if (!IsValueValidForKind(txt, ct.Kind))
                    {
                        MessageBox.Show($"[{name}] 열은 {ct.Kind} 형식입니다. 올바른 값을 입력하세요.", "XQLite",
                                         MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        // 되돌리기(선택): cell.Value2 = null;
                    }
                }
                finally { XqlCommon.ReleaseCom(lc); XqlCommon.ReleaseCom(h); XqlCommon.ReleaseCom(cell); }
            }
        }

        private static bool IsValueValidForKind(string? v, ColumnKind kind)
        {
            if (string.IsNullOrEmpty(v)) return true; // 빈칸 허용
            switch (kind)
            {
                case ColumnKind.Int: return long.TryParse(v, out _);
                case ColumnKind.Real: return double.TryParse(v, out _);
                case ColumnKind.Bool: return string.Equals(v, "TRUE", true) || string.Equals(v, "FALSE", true);
                case ColumnKind.Json:
                    try { Newtonsoft.Json.Linq.JToken.Parse(v); return true; } catch { return false; }
                default: return true;
            }
        }
#endif
    }
}
