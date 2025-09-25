// XqlSheetView.cs
using System;
using System.Collections.Generic;
using System.Linq;
using ExcelDna.Integration;
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
        private const string MetaMarkerName = "_XQL_META_HEADER"; // 시트당 1개

        // ───────────────────────── Public API (Ribbon에서 호출)
        public static void InsertMetaHeaderFromSelection()
        {
            var app = (Excel.Application)ExcelDnaUtil.Application;
            Excel.Worksheet? ws = null; Excel.Range? sel = null;
            try
            {
                ws = (Excel.Worksheet)app.ActiveSheet;
                sel = GetSelection(ws);
                if (ws == null) return;

                var sheet = XqlAddIn.Sheet;
                if (sheet == null) { MessageBox.Show("Sheet service not ready.", Caption); return; }

                var oldHeader = TryGetMarkerRange(ws);
                var newHeader = ResolveHeader(ws, sel, sheet) ?? XqlSheet.GetHeaderRange(ws);

                if (!SameAddress(oldHeader, newHeader))
                    ClearHeaderUi(ws, oldHeader); // 이전 UI 정리

                var names = BuildHeaderNames(newHeader);
                var sm = sheet.GetOrCreateSheet(ws.Name);
                sheet.EnsureColumns(ws.Name, names);
                var dict = sheet.BuildTooltipsForSheet(ws.Name);

                SetHeaderTooltips(newHeader, dict);
                ApplyHeaderOutlineBorder(newHeader);
                ApplyDataValidationForHeader(ws, newHeader, sm, sheet);

                SetMarker(ws, newHeader); // 단 하나
            }
            catch (Exception ex)
            {
                MessageBox.Show("InsertMetaHeader failed: " + ex.Message, Caption);
            }
            finally { XqlCommon.ReleaseCom(sel); XqlCommon.ReleaseCom(ws); }
        }

        public static void RefreshMetaHeader()
        {
            var app = (Excel.Application)ExcelDnaUtil.Application;
            Excel.Worksheet? ws = null;
            try
            {
                ws = (Excel.Worksheet)app.ActiveSheet;
                if (ws == null) return;

                var sheet = XqlAddIn.Sheet;
                if (sheet == null) { MessageBox.Show("Sheet service not ready.", Caption); return; }

                if (!sheet.TryGetSheet(ws.Name, out var sm))
                {
                    ClearHeaderUi(ws, TryGetMarkerRange(ws), removeMarker: true);
                    return;
                }

                var header = TryGetMarkerRange(ws) ?? ResolveHeader(ws, GetSelection(ws), sheet) ?? XqlSheet.GetHeaderRange(ws);
                var dict = sheet.BuildTooltipsForSheet(ws.Name);

                SetHeaderTooltips(header, dict);
                ApplyHeaderOutlineBorder(header);
                ApplyDataValidationForHeader(ws, header, sm, sheet);

                SetMarker(ws, header);
            }
            catch (Exception ex)
            {
                MessageBox.Show("RefreshMetaHeader failed: " + ex.Message, Caption);
            }
            finally { XqlCommon.ReleaseCom(ws); }
        }

        public static void RemoveMetaHeader()
        {
            var app = (Excel.Application)ExcelDnaUtil.Application;
            Excel.Worksheet? ws = null;
            try
            {
                ws = (Excel.Worksheet)app.ActiveSheet;
                if (ws == null) return;
                ClearHeaderUi(ws, TryGetMarkerRange(ws), removeMarker: true);
            }
            catch (Exception ex)
            {
                MessageBox.Show("RemoveMetaHeader failed: " + ex.Message, Caption);
            }
            finally { XqlCommon.ReleaseCom(ws); }
        }

        public static void ShowMetaHeaderInfo()
        {
            var app = (Excel.Application)ExcelDnaUtil.Application;
            Excel.Worksheet? ws = null; Excel.Range? sel = null;
            try
            {
                ws = (Excel.Worksheet)app.ActiveSheet;
                sel = GetSelection(ws);
                if (ws == null) return;

                var sheet = XqlAddIn.Sheet;
                if (sheet == null) { MessageBox.Show("Sheet service not ready.", Caption); return; }
                if (!sheet.TryGetSheet(ws.Name, out var sm))
                {
                    MessageBox.Show("No sheet meta found.", Caption);
                    return;
                }

                var header = TryGetMarkerRange(ws) ?? ResolveHeader(ws, sel, sheet) ?? XqlSheet.GetHeaderRange(ws);
                Excel.Range? hit = null, cell = null;
                try
                {
                    hit = (sel != null) ? XqlCommon.IntersectSafe(ws, header, sel) : null;
                    cell = (Excel.Range)((hit != null && hit.Cells.Count >= 1) ? hit.Cells[1, 1] : header.Cells[1, 1]);

                    var colName = (cell.Value2 as string)?.Trim();
                    if (string.IsNullOrEmpty(colName))
                        colName = XqlCommon.ColumnIndexToLetter(cell.Column);

                    if (!string.IsNullOrEmpty(colName) && sm.Columns.TryGetValue(colName!, out var ct))
                    {
                        MessageBox.Show($"{ws.Name}.{colName}\r\n{ct.ToTooltip()}", Caption);
                        return;
                    }
                }
                finally { XqlCommon.ReleaseCom(cell); XqlCommon.ReleaseCom(hit); }

                var lines = sm.Columns.OrderBy(kv => kv.Key, StringComparer.Ordinal)
                                      .Select(kv => $"{kv.Key} : {kv.Value.ToTooltip()}");
                var txt = string.Join("\r\n", lines);
                MessageBox.Show($"{ws.Name}\r\n{(string.IsNullOrEmpty(txt) ? "(no columns)" : txt)}", Caption);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Meta info failed: " + ex.Message, Caption);
            }
            finally { XqlCommon.ReleaseCom(sel); XqlCommon.ReleaseCom(ws); }
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

        private static Excel.Range? TryGetMarkerRange(Excel.Worksheet ws)
        {
            if (ws == null) return null;
            Excel.Names? names = null; Excel.Name? n = null; Excel.Range? r = null;
            try
            {
                names = ws.Names;
                try { n = names.Item(MetaMarkerName); } catch { n = null; }
                if (n == null) return null;

                try { r = n.RefersToRange; } catch { r = null; }
                return r;
            }
            catch { return null; }
            finally { XqlCommon.ReleaseCom(n); XqlCommon.ReleaseCom(names); }
        }

        private static void SetMarker(Excel.Worksheet ws, Excel.Range header)
        {
            if (ws == null || header == null) return;
            try { ws.Names.Item(MetaMarkerName)?.Delete(); } catch { }
            ws.Names.Add(Name: MetaMarkerName, RefersTo: header);
            try { ws.Names.Item(MetaMarkerName).Visible = false; } catch { }
        }

        private static void ClearMarker(Excel.Worksheet ws)
        {
            try { ws?.Names.Item(MetaMarkerName)?.Delete(); } catch { }
        }

        private static bool SameAddress(Excel.Range? a, Excel.Range? b)
        {
            try
            {
                if (a == null || b == null) return false;
                return string.Equals(a.Address[false, false], b.Address[false, false], StringComparison.Ordinal);
            }
            catch { return false; }
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

        internal static void SetHeaderTooltips(Excel.Range header, IReadOnlyDictionary<string, string> colToTip)
        {
            foreach (Excel.Range cell in header.Cells)
            {
                try
                {
                    var key = (cell.Value2 as string)?.Trim();
                    if (string.IsNullOrEmpty(key))
                        key = XqlCommon.ColumnIndexToLetter(cell.Column);

                    if (!colToTip.TryGetValue(key!, out var tip)) continue;

                    var c = cell.Comment;
                    try { c?.Delete(); }
                    catch { }
                    finally { XqlCommon.ReleaseCom(c); }

                    try { c = cell.AddComment(tip); }
                    catch { }
                    finally { XqlCommon.ReleaseCom(c); }
                }
                finally { XqlCommon.ReleaseCom(cell); }
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
                    try { b.LineStyle = Excel.XlLineStyle.xlContinuous; b.Weight = Excel.XlBorderWeight.xlMedium; }
                    finally { XqlCommon.ReleaseCom(b); }
                }
            }
            catch { }
            finally { XqlCommon.ReleaseCom(bs); }
        }

        private static void ClearHeaderUi(Excel.Worksheet ws, Excel.Range? header, bool removeMarker = false)
        {
            if (header == null) { if (removeMarker) ClearMarker(ws); return; }

            foreach (Excel.Range cell in header.Cells)
            {
                try
                {
                    var c = cell.Comment;
                    try { c?.Delete(); }
                    catch { }
                    finally { XqlCommon.ReleaseCom(c); }
                }
                finally { XqlCommon.ReleaseCom(cell); }
            }

            // Borders off
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
                    try { b.LineStyle = Excel.XlLineStyle.xlLineStyleNone; }
                    finally { XqlCommon.ReleaseCom(b); }
                }
            }
            catch { }
            finally { XqlCommon.ReleaseCom(bs); }

            if (removeMarker) ClearMarker(ws);
        }

        private static void ApplyDataValidationForHeader(Excel.Worksheet ws, Excel.Range header, SheetMeta sm, XqlSheet sheet)
        {
            var lo = XqlSheet.FindListObjectContaining(ws, header);
            if (lo?.HeaderRowRange != null)
            {
                for (int i = 1; i <= header.Columns.Count; i++)
                {
                    Excel.Range? h = null; Excel.Range? body = null;
                    try
                    {
                        h = (Excel.Range)header.Cells[1, i];
                        string? name = (h.Value2 as string)?.Trim();
                        if (string.IsNullOrEmpty(name)) 
                            name = XqlCommon.ColumnIndexToLetter(h.Column);

                        if (!sm.Columns.TryGetValue(name!, out var ct)) continue;

                        var lcol = lo.ListColumns[i];
                        body = lcol?.DataBodyRange;
                        if (body == null) continue;
                        ApplyValidationForKind(body, ct.Kind);
                    }
                    catch { }
                    finally { XqlCommon.ReleaseCom(h); XqlCommon.ReleaseCom(body); }
                }
                return;
            }

            // 일반 범위
            Excel.Range? used = null;
            try
            {
                used = ws.UsedRange;
                int lastRow = used.Row + used.Rows.Count - 1;
                int startRow = header.Row + 1;
                if (lastRow < startRow) return;

                for (int i = 1; i <= header.Columns.Count; i++)
                {
                    Excel.Range? h = null; Excel.Range? colRange = null;
                    try
                    {
                        h = (Excel.Range)header.Cells[1, i];
                        string? name = (h.Value2 as string)?.Trim();
                        if (string.IsNullOrEmpty(name)) 
                            name = XqlCommon.ColumnIndexToLetter(h.Column);

                        if (!sm.Columns.TryGetValue(name!, out var ct)) 
                            continue;

                        int col = h.Column;
                        colRange = ws.Range[ws.Cells[startRow, col], ws.Cells[lastRow, col]];
                        ApplyValidationForKind(colRange, ct.Kind);
                    }
                    catch { }
                    finally { XqlCommon.ReleaseCom(h); XqlCommon.ReleaseCom(colRange); }
                }
            }
            catch { }
            finally { XqlCommon.ReleaseCom(used); }
        }

        private static void ApplyValidationForKind(Excel.Range rng, ColumnKind kind)
        {
            try
            {
                try { rng.Validation?.Delete(); } catch { }

                switch (kind)
                {
                    case ColumnKind.Int:
                        rng.Validation!.Add(
                            Excel.XlDVType.xlValidateWholeNumber,
                            Excel.XlDVAlertStyle.xlValidAlertStop,
                            Excel.XlFormatConditionOperator.xlBetween,
                            -2147483648, 2147483647);
                        rng.Validation.IgnoreBlank = true;
                        rng.Validation.ErrorTitle = "정수만 허용";
                        rng.Validation.ErrorMessage = "이 열은 정수만 입력할 수 있습니다.";
                        break;

                    case ColumnKind.Real:
                        rng.Validation!.Add(
                            Excel.XlDVType.xlValidateDecimal,
                            Excel.XlDVAlertStyle.xlValidAlertStop,
                            Excel.XlFormatConditionOperator.xlBetween,
                            -1e307, 1e307);
                        rng.Validation.IgnoreBlank = true;
                        rng.Validation.ErrorTitle = "숫자만 허용";
                        rng.Validation.ErrorMessage = "이 열은 숫자(실수)만 입력할 수 있습니다.";
                        break;

                    case ColumnKind.Date:
                        rng.Validation!.Add(
                            Excel.XlDVType.xlValidateDate,
                            Excel.XlDVAlertStyle.xlValidAlertStop,
                            Excel.XlFormatConditionOperator.xlBetween,
                            new DateTime(1900, 1, 1), new DateTime(9999, 12, 31));
                        rng.Validation.IgnoreBlank = true;
                        rng.Validation.ErrorTitle = "날짜만 허용";
                        rng.Validation.ErrorMessage = "이 열은 날짜만 입력할 수 있습니다.";
                        break;

                    case ColumnKind.Bool:
                        rng.Validation!.Add(
                            Excel.XlDVType.xlValidateList,
                            Excel.XlDVAlertStyle.xlValidAlertStop,
                            Excel.XlFormatConditionOperator.xlBetween,
                            "TRUE,FALSE");
                        rng.Validation.IgnoreBlank = true;
                        rng.Validation.ErrorTitle = "TRUE/FALSE만 허용";
                        rng.Validation.ErrorMessage = "이 열은 TRUE 또는 FALSE만 입력할 수 있습니다.";
                        break;

                    case ColumnKind.Text:
                    default:
                        // 제한 없음
                        break;
                }
            }
            catch { /* 일부 병합/빈 범위에서 실패 가능 */ }
        }
    }
}
