// XqlSheetView.cs
using ExcelDna.Integration;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;
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

        // ── [NEW] 지연 재적용(QueueAsMacro) + 시트별 중복 큐잉 제거
        private static readonly object _reapplyLock = new object();
        private static readonly HashSet<string> _reapplyPending = new HashSet<string>(StringComparer.Ordinal);


        // ───────────────────────── Public API (Ribbon에서 호출)
        public static bool InstallHeader()
        {
            var app = (Excel.Application)ExcelDnaUtil.Application;
            Excel.Worksheet? ws = null; Excel.Range? candidate = null;
            try
            {
                ws = (Excel.Worksheet)app.ActiveSheet;
                if (ws == null) return false;

                var sheet = XqlAddIn.Sheet;
                if (sheet == null) { MessageBox.Show("Sheet service not ready.", "XQLite"); return false; }

                // ✅ 규칙: 시트당 헤더는 반드시 1개 — 기존 마커가 있으면 무조건 차단
                if (XqlSheet.TryGetHeaderMarker(ws, out var any))
                {
                    XqlCommon.ReleaseCom(any);
                    MessageBox.Show("이미 이 시트에는 헤더가 설치되어 있습니다.\r\n헤더를 제거한 뒤 다시 시도하세요.",
                    "XQLite", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return false;
                }

                // (헤더가 없을 때만 후보를 계산)
                candidate = GetHeaderOrFallback(ws);
                if (candidate == null)
                {
                    MessageBox.Show("헤더 후보를 찾을 수 없습니다.", "XQLite",
                                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return false;
                }

                // 메타 동기화
                var names = BuildHeaderNames(candidate);
                var sm = sheet.GetOrCreateSheet(ws.Name);
                sheet.EnsureColumns(ws.Name, names);

                // UI/검증 한 번에
                ApplyHeaderUi(ws, candidate, sm, withValidation: true);

                // 마커 확정
                XqlSheet.SetHeaderMarker(ws, candidate);

                // Excel 내부 후처리 이후 유실 방지(지연 재적용·중복 큐잉 방지)
                EnqueueReapplyHeaderUi(ws.Name, withValidation: true);
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("InstallHeader failed: " + ex.Message, "XQLite");
                return false;
            }
            finally { XqlCommon.ReleaseCom(candidate); XqlCommon.ReleaseCom(ws); }
        }

        // 메타에 있으면 메타 기반, 없으면 폴백
        private static string ColumnTooltipFor(SheetMeta sm, string colName)
        {
            try
            {
                if (!string.IsNullOrWhiteSpace(colName) &&
                    sm.Columns != null &&
                    sm.Columns.TryGetValue(colName, out var ct) && ct != null)
                {
                    // SheetMeta.ColumnMeta가 ToTooltip()을 제공한다면 그대로 활용
                    try { return ct.ToTooltip(); } catch { /* fall through */ }
                    // 제공하지 않으면 최소 정보 구성 (Kind/Null/Check 등 프로젝트 모델에 맞춰 보강 가능)
                    return $"{ct.Kind} • {(ct.Nullable ? "NULL OK" : "NOT NULL")}";
                }
            }
            catch { /* ignore */ }

            return ColumnTooltipFallback(); // 폴백
        }

        private static string ColumnTooltipFallback() => "TEXT • NULL OK";

        // 헤더 범위를 돌며 i열(1-based) → 툴팁 텍스트를 만든다.
        internal static IReadOnlyDictionary<int, string> BuildHeaderTooltips(SheetMeta sm, Excel.Range header)
        {
            var tips = new Dictionary<int, string>(capacity: Math.Max(1, header?.Columns.Count ?? 0));
            if (header == null) return tips;

            int cols = header.Columns.Count;
            for (int i = 1; i <= cols; i++)
            {
                Excel.Range? h = null;
                try
                {
                    h = (Excel.Range)header.Cells[1, i];
                    var colName = (h.Value2 as string)?.Trim();
                    if (string.IsNullOrEmpty(colName))
                        colName = XqlCommon.ColumnIndexToLetter(h.Column);

                    tips[i] = ColumnTooltipFor(sm, colName!); // 아래 헬퍼 사용
                }
                finally { XqlCommon.ReleaseCom(h); }
            }
            return tips;
        }

        public static void RefreshHeader()
        {
            var app = (Excel.Application)ExcelDnaUtil.Application;
            Excel.Worksheet? ws = null; Excel.Range? header = null;
            try
            {
                ws = (Excel.Worksheet)app.ActiveSheet; if (ws == null) return;
                header = GetHeaderOrFallback(ws);
                if (header == null) { MessageBox.Show("헤더를 찾을 수 없습니다.", "XQLite"); return; }

                RebindMarkerIfMoved(ws, header);

                var sheet = XqlAddIn.Sheet!;
                var sm = sheet.GetOrCreateSheet(ws.Name);

                ApplyHeaderUi(ws, header, sm, withValidation: true);

                EnqueueReapplyHeaderUi(ws.Name, withValidation: true);
            }
            catch (Exception ex) { MessageBox.Show("RefreshMetaHeader failed: " + ex.Message, "XQLite"); }
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
            Excel.Worksheet? ws = null; Excel.Range? header = null;
            try
            {
                ws = (Excel.Worksheet)app.ActiveSheet; if (ws == null) return;
                header = GetHeaderOrFallback(ws);
                if (header == null) { MessageBox.Show("헤더를 찾을 수 없습니다.", "XQLite"); return; }

                // 이동했으면 마커 재바인딩
                RebindMarkerIfMoved(ws, header);

                var sheet = XqlAddIn.Sheet!;
                var sm = sheet.GetOrCreateSheet(ws.Name);

                var sb = new System.Text.StringBuilder();
                var addr = header.Address[true, true, Excel.XlReferenceStyle.xlA1, false];
                sb.AppendLine($"{ws.Name}!{addr}");
                sb.AppendLine($"Start: Col {XqlCommon.ColumnIndexToLetter(header.Column)} ({header.Column}), Row {header.Row}  |  Data @ {header.Row + 1}");
                sb.AppendLine();

                for (int i = 1; i <= header.Columns.Count; i++)
                {
                    Excel.Range? h = null;
                    try
                    {
                        h = (Excel.Range)header.Cells[1, i];
                        var name = (h.Value2 as string)?.Trim();
                        if (string.IsNullOrEmpty(name)) name = XqlCommon.ColumnIndexToLetter(h.Column);
                        if (sm.Columns.TryGetValue(name!, out var ct))
                            sb.AppendLine($"{ws.Name}.{name}\r\n{ct.ToTooltip()}");
                        else
                            sb.AppendLine($"{ws.Name}.{name}\r\nTEXT • NULL OK");
                        sb.AppendLine();
                    }
                    finally { XqlCommon.ReleaseCom(h); }
                }
                MessageBox.Show(sb.ToString().TrimEnd(), "XQLite");
            }
            catch (Exception ex) { MessageBox.Show("ShowMetaHeaderInfo failed: " + ex.Message, "XQLite"); }
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


        private static string SafeCommentText(Excel.Comment c)
        {
            try { return (c.Text() ?? "").Trim(); } catch { return ""; }
        }

        // 헤더 코멘트(툴팁)를 깔끔하게 갱신: 같으면 건너뛰고, 다르면 '제자리(Text) 갱신' 우선
        internal static void SetHeaderTooltips(Excel.Range header, IReadOnlyDictionary<int, string> tips)
        {
            if (header == null || tips == null || tips.Count == 0) return;

            var app = (Excel.Application)header.Application;
            bool oldSU = true, oldEv = true;
            try
            {
                // 화면/이벤트 잠시 OFF → 플리커 최소화
                try { oldSU = app.ScreenUpdating; app.ScreenUpdating = false; } catch { }
                try { oldEv = app.EnableEvents; app.EnableEvents = false; } catch { }

                int cols = header.Columns.Count;
                for (int i = 1; i <= cols; i++)
                {
                    Excel.Range? cell = null; Excel.Comment? cmt = null;
                    try
                    {
                        cell = (Excel.Range)header.Cells[1, i];
                        var text = tips.TryGetValue(i, out var t) ? t : string.Empty;
                        if (!string.IsNullOrEmpty(text) && text.Length > 512)
                            text = text.Substring(0, 509) + "...";

                        // 비워야 하면 삭제(이미 없으면 skip)
                        if (string.IsNullOrEmpty(text))
                        {
                            try { cell.Comment?.Delete(); } catch { }
                            continue;
                        }

                        cmt = cell.Comment;
                        if (cmt != null)
                        {
                            // 같다면 아무 것도 안 함(재그림 없음)
                            var cur = SafeCommentText(cmt);
                            if (string.Equals(cur, text, StringComparison.Ordinal)) continue;

                            // 제자리 갱신 시도 → 실패하면 최후에 delete+add
                            try { cmt.Text(text); }
                            catch
                            {
                                try { cmt.Delete(); } catch { }
                                try { cell.AddComment(text); } catch { /* ignore */ }
                            }
                        }
                        else
                        {
                            // 없을 때만 Add(삼각형 신규 생성) → 플리커 횟수 최소화
                            try { cell.AddComment(text); } catch { /* ignore */ }
                        }
                    }
                    finally { XqlCommon.ReleaseCom(cmt); XqlCommon.ReleaseCom(cell); }
                }
            }
            finally
            {
                // 원복은 실패해도 무시
                try { app.EnableEvents = oldEv; } catch { }
                try { app.ScreenUpdating = oldSU; } catch { }
            }
        }


        // 헤더 1행이 편집되면 툴팁 재적용
        public static void RefreshTooltipsIfHeaderEdited(Excel.Worksheet ws, Excel.Range target)
        {
            if (ws == null || target == null) return;
            var sheetSvc = XqlAddIn.Sheet; if (sheetSvc == null) return;

            Excel.Range? marker = null, inter = null;
            Excel.ListObject? lo = null;
            bool isHeaderEdit = false;
            string sheetName = ws.Name;

            try
            {
                // 1) 마커 기준 교차 검사
                if (XqlSheet.TryGetHeaderMarker(ws, out marker))
                {
                    inter = ws.Application.Intersect(marker, target);
                    isHeaderEdit = inter != null;
                }

                // 2) 마커가 없거나 교차 안되면, 표 헤더 교차로 한 번 더 확인
                if (!isHeaderEdit)
                {
                    lo = XqlSheet.FindListObjectContaining(ws, target);
                    var hdr = lo?.HeaderRowRange;
                    if (hdr != null)
                    {
                        var hit = XqlCommon.IntersectSafe(ws, hdr, target);
                        isHeaderEdit = hit != null;
                        XqlCommon.ReleaseCom(hit);
                    }
                }

                if (!isHeaderEdit) return;

                // 3) Excel이 헤더 갱신을 끝낸 이후에 재적용 (UI 스레드 매크로 큐)
                ExcelAsyncUtil.QueueAsMacro(() =>
                {
                    Excel.Worksheet? ws2 = null; Excel.Range? header2 = null;
                    try
                    {
                        var app2 = (Excel.Application)ExcelDnaUtil.Application;
                        // sheet 객체를 직접 들고오지 말고 이름으로 다시 획득 (COM 안정)
                        ws2 = XqlSheet.FindWorksheet(app2, sheetName);
                        if (ws2 == null) return;

                        // 새로 계산된 헤더 범위 확보 (마커 → 폴백 순)
                        if (!XqlSheet.TryGetHeaderMarker(ws2, out header2))
                            header2 = XqlSheet.GetHeaderRange(ws2);

                        var sm = sheetSvc.GetOrCreateSheet(sheetName);
                        ApplyHeaderUi(ws2, header2, sm, withValidation: true);
                    }
                    catch { /* 무음 */ }
                    finally { XqlCommon.ReleaseCom(header2); XqlCommon.ReleaseCom(ws2); }
                });
            }
            finally
            {
                XqlCommon.ReleaseCom(inter);
                XqlCommon.ReleaseCom(marker);
                XqlCommon.ReleaseCom(lo);
            }
        }

        internal static void ApplyDataValidationForHeader(Excel.Worksheet ws, Excel.Range header, SheetMeta sm)
        {
            var lo = XqlSheet.FindListObjectContaining(ws, header);
            if (lo?.HeaderRowRange != null)
            {
                for (int i = 1; i <= header.Columns.Count; i++)
                {
                    Excel.Range? h = null; Excel.Range? rng = null;
                    try
                    {
                        h = (Excel.Range)header.Cells[1, i];
                        string? name = (h.Value2 as string)?.Trim();
                        if (string.IsNullOrEmpty(name)) name = XqlCommon.ColumnIndexToLetter(h.Column);

                        // 표가 있으면 그 컬럼의 DataBodyRange에만 DV 적용
                        try { rng = lo.ListColumns[i]?.DataBodyRange; } catch { rng = null; }
                        if (rng == null) rng = ColBelowToEnd(ws, h); // 표가 비어 있으면 폴백

                        if (sm.Columns.TryGetValue(name!, out var ct))
                            ApplyValidationForKind(rng, ct.Kind);
                        else
                            try { rng.Validation.Delete(); } catch { /* clean only */ }
                    }
                    finally { XqlCommon.ReleaseCom(h); XqlCommon.ReleaseCom(rng); }
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
                        try { col.Validation.Delete(); } catch { /* clean only */ }
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

        // Excel의 DV는 일부 타입(TEXT/JSON 등)엔 굳이 깔지 않는다.
        // 아래 로직은 '규칙을 실제로 Add한 경우에만' 속성을 세팅하여 0x800A03EC를 방지한다.
        private static void ApplyValidationForKind(Excel.Range rng, ColumnKind kind)
        {
            Excel.Validation? v = null;
            try
            {
                // 빈/다중 영역은 스킵 (Excel DV가 잘 안 먹거나 예외 발생)
                try
                {
                    if (rng == null) return;
                    if ((long)rng.CountLarge == 0) return;
                    if (rng.Areas != null && rng.Areas.Count > 1) return;
                }
                catch { /* ignore */ }

                // 기존 규칙 제거 (잔존/중복으로 Add 실패 방지)
                try { rng.Validation.Delete(); } catch { }

                bool added = false;
                v = rng.Validation;

                switch (kind)
                {
                    case ColumnKind.Int:
                        v.Add(
                            Excel.XlDVType.xlValidateWholeNumber,
                            Excel.XlDVAlertStyle.xlValidAlertStop,
                            Excel.XlFormatConditionOperator.xlBetween,
                            -2147483648, 2147483647);
                        v.IgnoreBlank = true;
                        v.ErrorTitle = "정수만 허용";
                        v.ErrorMessage = "이 열은 정수만 입력할 수 있습니다.";
                        added = true;
                        break;

                    case ColumnKind.Real:
                        v.Add(
                            Excel.XlDVType.xlValidateDecimal,
                            Excel.XlDVAlertStyle.xlValidAlertStop,
                            Excel.XlFormatConditionOperator.xlBetween,
                            -1.79e308, 1.79e308);
                        v.IgnoreBlank = true;
                        v.ErrorTitle = "실수만 허용";
                        v.ErrorMessage = "이 열은 실수/숫자만 입력할 수 있습니다.";
                        added = true;
                        break;

                    case ColumnKind.Bool:
                        // TRUE/FALSE 리스트로 제한
                        v.Add(
                            Excel.XlDVType.xlValidateList,
                            Excel.XlDVAlertStyle.xlValidAlertStop,
                            Type.Missing, "TRUE,FALSE");
                        v.IgnoreBlank = true;
                        v.ErrorTitle = "BOOL만 허용";
                        v.ErrorMessage = "이 열은 TRUE 또는 FALSE만 입력할 수 있습니다.";
                        added = true;
                        break;

                    case ColumnKind.Date:
                        // 날짜만: 1900-01-01 ~ 9999-12-31
                        v.Add(
                            Excel.XlDVType.xlValidateDate,
                            Excel.XlDVAlertStyle.xlValidAlertStop,
                            Excel.XlFormatConditionOperator.xlBetween,
                            "=DATE(1900,1,1)", "=DATE(9999,12,31)");
                        v.IgnoreBlank = true;
                        v.ErrorTitle = "날짜만 허용";
                        v.ErrorMessage = "이 열은 날짜 형식만 입력할 수 있습니다.";
                        added = true;
                        break;

                    // TEXT/JSON/ANY 등은 DV를 깔지 않는다(서버/런타임 검증으로 커버)
                    // case ColumnKind.Text:
                    // case ColumnKind.Json:
                    default:
                        added = false;
                        break;
                }

                // ✨ DV를 실제로 추가한 경우에만 속성 설정 → 0x800A03EC 방지
                if (added)
                {
                    try { v.ShowError = true; } catch { }
                    try { v.ShowInput = false; } catch { }
                }
            }
            catch
            {
                // 병합/보호/특수범위 등으로 Add가 실패할 수 있음 — 침묵
            }
            finally
            {
                XqlCommon.ReleaseCom(v);
            }
        }


        // 헤더: 마커 → (선택 기반) ResolveHeader → Fallback(GetHeaderRange) 순서로 결정
        private static Excel.Range? GetHeaderOrFallback(Excel.Worksheet ws)
        {
            if (XqlSheet.TryGetHeaderMarker(ws, out var hdr)) return hdr;
            Excel.Range? sel = null; Excel.Range? guess = null;
            try
            {
                sel = GetSelection(ws);
                guess = ResolveHeader(ws, sel, XqlAddIn.Sheet!) ?? XqlSheet.GetHeaderRange(ws);
                return guess;
            }
            finally { XqlCommon.ReleaseCom(sel); /* guess는 반환 */ }
        }

        // 선택/추정된 헤더가 실제로 이동된 경우 마커를 새 위치로 동기화
        private static void RebindMarkerIfMoved(Excel.Worksheet ws, Excel.Range candidate)
        {
            if (XqlSheet.TryGetHeaderMarker(ws, out var old))
            {
                try { if (!XqlSheet.IsSameRange(old, candidate)) XqlSheet.SetHeaderMarker(ws, candidate); }
                finally { XqlCommon.ReleaseCom(old); }
            }
            else
            {
                XqlSheet.SetHeaderMarker(ws, candidate);
            }
        }

        // ── [NEW] 헤더 UI(툴팁+보더+검증) 한 번에 적용
        private static void ApplyHeaderUi(Excel.Worksheet ws, Excel.Range header, SheetMeta sm, bool withValidation)
        {
            if (ws == null || header == null || sm == null) return;

            // 툴팁 + 보더
            var tips = BuildHeaderTooltips(sm, header);
            SetHeaderTooltips(header, tips);
            ApplyHeaderOutlineBorder(header);

            // 데이터 검증(옵션): 열 끝까지, 표 유무 무관
            if (withValidation)
                ApplyDataValidationForHeader(ws, header, sm);
        }

        // Excel 내부가 헤더를 재구성하는 타이밍을 기다렸다가 재적용(디바운스)
        private static void EnqueueReapplyHeaderUi(string sheetName, bool withValidation)
        {
            string key = $"{sheetName}:{withValidation}";
            lock (_reapplyLock)
            {
                if (!_reapplyPending.Add(key)) return; // 이미 대기 중이면 무시
            }

            // 약간 기다렸다가 메인 스레드에서 일괄 재적용 → 깜빡임 최소
            Task.Run(async () =>
            {
                await Task.Delay(120).ConfigureAwait(false); // 80~150ms 권장

                ExcelAsyncUtil.QueueAsMacro(() =>
                {
                    Excel.Worksheet? ws2 = null; Excel.Range? h2 = null;
                    try
                    {
                        var app2 = (Excel.Application)ExcelDnaUtil.Application;
                        ws2 = XqlSheet.FindWorksheet(app2, sheetName);
                        if (ws2 == null) return;

                        if (!XqlSheet.TryGetHeaderMarker(ws2, out h2))
                            h2 = XqlSheet.GetHeaderRange(ws2);
                        if (h2 == null) return;

                        var sm = XqlAddIn.Sheet!.GetOrCreateSheet(sheetName);
                        ApplyHeaderUi(ws2, h2, sm, withValidation); // 툴팁+보더(+검증)
                    }
                    finally
                    {
                        XqlCommon.ReleaseCom(h2);
                        XqlCommon.ReleaseCom(ws2);
                        lock (_reapplyLock) { _reapplyPending.Remove(key); }
                    }
                });
            });
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
