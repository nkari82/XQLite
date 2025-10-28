// XqlSheetView.cs
using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using static XQLite.AddIn.XqlCommon;
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

        private static readonly object _sumLock = new();
        private static HashSet<string> _sumTables = new(StringComparer.Ordinal);
        private static int _sumAffected, _sumConflicts, _sumErrors, _sumBatches;
        private static long _sumStartTicks;

        // 문자열/맵만 캐시(⚠ COM 미보관)
        private static readonly ConcurrentDictionary<string, string> _tableToSheet = new(StringComparer.Ordinal);
        private static readonly ConcurrentDictionary<string, (string addr, Dictionary<string, string> map)> _hdrCache = new(StringComparer.Ordinal);

        // ───────────────────────── 공통 파이프라인
        private enum HeaderOp { Install, Refresh, Remove }

        // 스펙: 헤더를 지정한 순서로 강제 정렬하고 싶은 경우 사용 (null이면 현 상태 유지)
        private readonly struct HeaderSpec
        {
            public readonly IList<string> Columns;
            public HeaderSpec(IList<string> columns) { Columns = columns; }
        }

        private static bool RunHeaderPipeline(HeaderOp op)
        {
            var app = (Excel.Application)ExcelDnaUtil.Application;
            using var wsW = SmartCom<Worksheet>.Wrap((Excel.Worksheet)app.ActiveSheet);
            if (wsW.Value == null) return false;

            try
            {
                return ExecuteHeaderPipeline(app, wsW.Value!, op, null);
            }
            catch (Exception ex)
            {
                var msg = op switch
                {
                    HeaderOp.Install => "InstallHeader failed: ",
                    HeaderOp.Refresh => "RefreshHeader failed: ",
                    _ => "RemoveHeader failed: "
                };
                MessageBox.Show(msg + ex.Message, Caption);
                return false;
            }
        }

        // 스펙을 받는 오버로드 (필요 시 직접 사용)
        private static bool RunHeaderPipeline(HeaderOp op, HeaderSpec? spec)
        {
            var app = (Excel.Application)ExcelDnaUtil.Application;
            using var wsW = SmartCom<Worksheet>.Wrap((Excel.Worksheet)app.ActiveSheet);
            if (wsW.Value == null) return false;

            try
            {
                return ExecuteHeaderPipeline(app, wsW.Value!, op, spec);
            }
            catch (Exception ex)
            {
                var msg = op switch
                {
                    HeaderOp.Install => "InstallHeader failed: ",
                    HeaderOp.Refresh => "RefreshHeader failed: ",
                    _ => "RemoveHeader failed: "
                };
                MessageBox.Show(msg + ex.Message, Caption);
                return false;
            }
        }

        // 스펙을 받아서 컬럼 순서를 “강제”할 수도 있는 본체
        private static bool ExecuteHeaderPipeline(Excel.Application app, Excel.Worksheet ws, HeaderOp op, HeaderSpec? spec)
        {
            // Remove는 별도 간단 경로
            if (op == HeaderOp.Remove)
            {
                Excel.Range? hdr = null;

                // 마커가 있으면 그것, 없으면 선택/폴백으로 해석
                if (!XqlSheet.TryGetHeaderMarker(ws, out hdr))
                {
                    using var selW = SmartCom<Range>.Wrap(GetSelection(ws));
                    hdr = ResolveHeader(ws, selW.Value, XqlAddIn.Sheet!) ?? XqlSheet.GetHeaderRange(ws);
                }

                ClearHeaderUi(ws, hdr, removeMarker: true);
                InvalidateHeaderCache(ws.Name);
                return true;
            }

            // Install / Refresh 공통 처리
            var sheetSvc = XqlAddIn.Sheet;
            if (sheetSvc == null)
            {
                MessageBox.Show("Sheet service not ready.", Caption);
                return false;
            }

            // Install 시 이미 설치된 헤더가 있으면 중단
            if (op == HeaderOp.Install && XqlSheet.TryGetHeaderMarker(ws, out var any))
            {
                using var _any = SmartCom<Range>.Wrap(any);
                MessageBox.Show("이미 이 시트에는 헤더가 설치되어 있습니다.\r\n헤더를 제거한 뒤 다시 시도하세요.",
                    Caption, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }

            // Install은 설치 후보(선택 1행 우선), Refresh는 마커/폴백
            Excel.Range? hdrCandidate = op == HeaderOp.Install
                ? GetHeaderCandidateForInstall(ws)
                : GetHeaderOrFallback(ws);

            using var hdrW = SmartCom<Range>.Wrap(hdrCandidate);
            if (hdrW.Value == null)
            {
                MessageBox.Show("헤더를 찾을 수 없습니다.", Caption, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }

            // 이동했으면 마커 재바인딩(Refresh/재설치 후 일관)
            RebindMarkerIfMoved(ws, hdrW.Value!);

            var sm = sheetSvc.GetOrCreateSheet(ws.Name);
            if (string.IsNullOrWhiteSpace(sm.KeyColumn)) sm.KeyColumn = "id";
            var keyName = sm.KeyColumn!;

            using (new XqlCommon.ExcelBatchScope(app))
            {
                Excel.Range newHeaderRaw;
                List<string> names;

                // ✅ 스펙이 있으면 그 순서/구성으로 “강제 정렬”
                if (spec is HeaderSpec s && s.Columns != null && s.Columns.Count > 0)
                {
                    newHeaderRaw = UpdateHeaderToColumns(ws, hdrW.Value!, sm, ws.Name, s.Columns);
                    names = s.Columns.ToList();
                }
                else
                {
                    // ✅ 아니면 현 헤더를 정리: id 1열 보정 + 이름 재구성
                    (newHeaderRaw, names) = EnsureIdFirstAndRebuildHeader(ws, hdrW.Value!, keyName);
                    sheetSvc.EnsureColumns(ws.Name, names);
                }

                using var newHeaderW = SmartCom<Range>.Wrap(newHeaderRaw);

                // 메타 컬럼 보장 + UI 적용 + 마커
                ApplyHeaderUi(ws, newHeaderW.Value!, sm, withValidation: true);
                XqlSheet.SetHeaderMarker(ws, newHeaderW.Value!);
            }

            InvalidateHeaderCache(ws.Name);
            return true;
        }

        // ───────────────────────── Public API (Ribbon에서 호출) — 얇은 래퍼만 남김
        public static bool InstallHeader() => RunHeaderPipeline(HeaderOp.Install);
        public static void RefreshHeader() => RunHeaderPipeline(HeaderOp.Refresh);
        public static void RemoveHeader() => RunHeaderPipeline(HeaderOp.Remove);

        // ───────────────────────── 기존/보조 로직 (그대로 유지 또는 안전성 보강)

        // 메타에 있으면 메타 기반, 없으면 폴백
        private static string ColumnTooltipFor(XqlSheet.Meta sm, string colName)
        {
            try
            {
                if (string.Equals(colName, sm.KeyColumn, StringComparison.OrdinalIgnoreCase))
                    return "INTEGER • NOT NULL • PRIMARY KEY";

                if (!string.IsNullOrWhiteSpace(colName) &&
                    sm.Columns != null &&
                    sm.Columns.TryGetValue(colName, out var ct) && ct != null)
                {
                    try { return ct.ToTooltip(); } catch { /* fall through */ }
                    return $"{ct.Kind} • {(ct.Nullable ? "NULL OK" : "NOT NULL")}";
                }
            }
            catch { /* ignore */ }

            return ColumnTooltipFallback();
        }

        private static string ColumnTooltipFallback() => "TEXT • NULL OK";

        internal static IReadOnlyDictionary<int, string> BuildHeaderTooltips(XqlSheet.Meta sm, Excel.Range header)
        {
            int capacity = 1;
            try { capacity = Math.Max(1, header?.Columns.Count ?? 0); } catch { capacity = 1; }

            var tips = new Dictionary<int, string>(capacity);
            if (header == null) return tips;

            int cols = 0;
            try { cols = header.Columns.Count; } catch { cols = 0; }
            for (int i = 1; i <= cols; i++)
            {
                using var h = SmartCom<Range>.Wrap((Excel.Range)header.Cells[1, i]);
                string? colName = null;
                try { colName = (h.Value?.Value2 as string)?.Trim(); } catch { colName = null; }
                if (string.IsNullOrEmpty(colName))
                {
                    try { colName = XqlCommon.ColumnIndexToLetter(h.Value!.Column); } catch { colName = null; }
                }
                tips[i] = ColumnTooltipFor(sm, colName ?? "");
            }
            return tips;
        }

        public static void ShowHeaderInfo()
        {
            var app = (Excel.Application)ExcelDnaUtil.Application;
            using var wsW = SmartCom<Worksheet>.Wrap((Excel.Worksheet)app.ActiveSheet);
            if (wsW.Value == null) return;

            using var headerW = SmartCom<Range>.Wrap(GetHeaderOrFallback(wsW.Value));
            if (headerW.Value == null) { MessageBox.Show("헤더를 찾을 수 없습니다.", "XQLite"); return; }

            try
            {
                RebindMarkerIfMoved(wsW.Value, headerW.Value);

                var sheet = XqlAddIn.Sheet!;
                var sm = sheet.GetOrCreateSheet(wsW.Value.Name);

                var sb = new System.Text.StringBuilder();
                var addr = headerW.Value.Address[true, true, Excel.XlReferenceStyle.xlA1, false];
                sb.AppendLine($"{wsW.Value.Name}!{addr}");
                sb.AppendLine($"Start: Col {XqlCommon.ColumnIndexToLetter(headerW.Value.Column)} ({headerW.Value.Column}), Row {headerW.Value.Row}  |  Data @ {headerW.Value.Row + 1}");
                sb.AppendLine();

                int cols = 0;
                try { cols = headerW.Value.Columns.Count; } catch { cols = 0; }

                for (int i = 1; i <= cols; i++)
                {
                    using var h = SmartCom<Range>.Wrap((Excel.Range)headerW.Value.Cells[1, i]);
                    var name = (h.Value?.Value2 as string)?.Trim();
                    if (string.IsNullOrEmpty(name)) name = XqlCommon.ColumnIndexToLetter(h.Value!.Column);

                    if (sm.Columns.TryGetValue(name!, out var ct))
                        sb.AppendLine($"{wsW.Value.Name}.{name}\r\n{ct.ToTooltip()}");
                    else
                        sb.AppendLine($"{wsW.Value.Name}.{name}\r\nTEXT • NULL OK");

                    sb.AppendLine();
                }
                MessageBox.Show(sb.ToString().TrimEnd(), "XQLite");
            }
            catch (Exception ex) { MessageBox.Show("ShowMetaHeaderInfo failed: " + ex.Message, "XQLite"); }
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
                using var s = SmartCom<Range>.Wrap(ws.Cells[r, c1] as Excel.Range);
                using var e = SmartCom<Range>.Wrap(ws.Cells[r, c2] as Excel.Range);
                return SmartCom<Range>.Wrap(ws.Range[s.Value, e.Value]).Detach();
            }
            return null;
        }

        public static bool TryGetHeaderSelectedColumn(Excel.Worksheet ws, out Excel.Range? headerCell, out string? colName)
        {
            headerCell = null; colName = null;
            if (ws == null) return false;

            var hdr = GetHeaderOrFallback(ws);
            if (hdr == null) return false;

            Excel.Range? sel = null;
            try { sel = (Excel.Range)ws.Application.Selection; } catch { }
            if (sel == null) return false;

            var inter = XqlCommon.IntersectSafe(ws, hdr, sel);
            if (inter == null) return false;

            int cells = 1;
            try { cells = inter.Cells.Count; } catch { }
            if (cells != 1) return false;

            Excel.Range cell;
            try { cell = (Excel.Range)inter.Cells[1, 1]; } catch { return false; }

            string name = null!;
            try { name = Convert.ToString(cell.Value2)?.Trim() ?? ""; } catch { }
            if (string.IsNullOrEmpty(name))
            {
                try { name = XqlCommon.ColumnIndexToLetter(cell.Column); } catch { name = ""; }
            }
            if (string.IsNullOrEmpty(name)) return false;

            headerCell = cell;
            colName = name;
            return true;
        }

        public static bool EnsureHeaderColumnSelectionOrWarn(Excel.Worksheet ws, string title = "XQLite")
        {
            if (TryGetHeaderSelectedColumn(ws, out _, out _))
                return true;

            try
            {
                MessageBox.Show("컬럼 타입은 '헤더의 한 셀'을 선택한 상태에서만 변경할 수 있습니다.\r\n" +
                                "헤더(1행)에서 해당 컬럼명을 클릭하고 다시 시도하세요.",
                                title, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            catch { /* ignore UI errors */ }
            return false;
        }

        // 설치 최초 후보(선택 1행 폭 유지 → 없으면 1행 기본)
        private static Excel.Range? GetHeaderCandidateForInstall(Excel.Worksheet ws)
        {
            if (ws == null) return null;
            if (XqlSheet.TryGetHeaderMarker(ws, out var hdr)) return hdr;

            var sel = GetSelection(ws);
            if (sel != null)
            {
                try
                {
                    int top = sel.Row;
                    int rows = sel.Rows.Count;
                    int cols = sel.Columns.Count;
                    if (top == 1 && rows == 1 && cols >= 1)
                    {
                        using var s = SmartCom<Range>.Wrap((Excel.Range)ws.Cells[1, sel.Column]);
                        using var e = SmartCom<Range>.Wrap((Excel.Range)ws.Cells[1, sel.Column + cols - 1]);
                        return SmartCom<Range>.Wrap(ws.Range[s.Value, e.Value]).Detach();
                    }
                }
                catch { /* ignore */ }
            }
            return XqlSheet.GetHeaderRange(ws);
        }

        private static Excel.Range? GetSelection(Excel.Worksheet ws)
        {
            try { return (Excel.Range)ws.Application.Selection; }
            catch { return null; }
        }

        // ───────────────────────── UI Helpers (Comments / Borders / Tooltips / Validation)

        private static List<string> BuildHeaderNames(Excel.Range header)
            => XqlSheet.ComputeHeaderNames(header);

        private static string SafeCommentText(Excel.Comment c)
        {
            try { return (c.Text() ?? "").Trim(); } catch { return ""; }
        }

        internal static void SetHeaderTooltips(Excel.Range header, IReadOnlyDictionary<int, string> tips)
        {
            if (header == null || tips == null || tips.Count == 0) return;

            var app = (Excel.Application)header.Application;
            bool oldSU = true, oldEv = true;
            try
            {
                try { oldSU = app.ScreenUpdating; app.ScreenUpdating = false; } catch { }
                try { oldEv = app.EnableEvents; app.EnableEvents = false; } catch { }

                int cols = 0;
                try { cols = header.Columns.Count; } catch { cols = 0; }

                for (int i = 1; i <= cols; i++)
                {
                    using var cell = SmartCom<Range>.Wrap((Excel.Range)header.Cells[1, i]);
                    Excel.Comment? cmt = null;
                    try
                    {
                        var text = tips.TryGetValue(i, out var t) ? t : string.Empty;
                        if (!string.IsNullOrEmpty(text) && text.Length > 512)
                            text = text.Substring(0, 509) + "...";

                        if (string.IsNullOrEmpty(text))
                        {
                            try { cell.Value?.Comment?.Delete(); } catch { }
                            continue;
                        }

                        cmt = cell.Value?.Comment;
                        if (cmt != null)
                        {
                            var cur = SafeCommentText(cmt);
                            if (string.Equals(cur, text, StringComparison.Ordinal)) continue;

                            try { cmt.Text(text); }
                            catch
                            {
                                try { cmt.Delete(); } catch { }
                                try { cell.Value?.AddComment(text); } catch { /* ignore */ }
                            }
                        }
                        else
                        {
                            try { cell.Value?.AddComment(text); } catch { /* ignore */ }
                        }
                    }
                    finally { /* RCW는 GC 위임 */ }
                }
            }
            finally
            {
                try { app.EnableEvents = oldEv; } catch { }
                try { app.ScreenUpdating = oldSU; } catch { }
            }
        }

        public static void RefreshTooltipsIfHeaderEdited(Excel.Worksheet ws, Excel.Range target)
        {
            if (ws == null || target == null) return;
            var sheetSvc = XqlAddIn.Sheet; if (sheetSvc == null) return;

            Excel.Range? marker = null;
            bool isHeaderEdit = false;
            string sheetName = ws.Name;

            try
            {
                if (XqlSheet.TryGetHeaderMarker(ws, out marker))
                {
                    using var inter = SmartCom<Range>.Wrap(ws.Application.Intersect(marker, target));
                    isHeaderEdit = inter.Value != null;
                }

                if (!isHeaderEdit)
                {
                    var lo = XqlSheet.FindListObjectContaining(ws, target);
                    var hdr = lo?.HeaderRowRange;
                    if (hdr != null)
                    {
                        using var hit = SmartCom<Range>.Wrap(XqlCommon.IntersectSafe(ws, hdr, target));
                        isHeaderEdit = hit.Value != null;
                    }
                }

                if (!isHeaderEdit) return;

                _ = XqlCommon.OnExcelThreadAsync(() =>
                {
                    var app2 = (Excel.Application)ExcelDnaUtil.Application;
                    using var ws2 = SmartCom<Worksheet>.Wrap(XqlSheet.FindWorksheet(app2, sheetName));
                    if (ws2.Value == null) return (object?)null;

                    using var header2 = SmartCom<Range>.Wrap(XqlSheet.TryGetHeaderMarker(ws2.Value, out var h) ? h : XqlSheet.GetHeaderRange(ws2.Value));
                    var sm = sheetSvc.GetOrCreateSheet(sheetName);
                    ApplyHeaderUi(ws2.Value, header2.Value!, sm, withValidation: true);
                    return (object?)null;
                });
            }
            finally { /* marker는 TryGetHeaderMarker 반환이므로 여기서 해제 금지 */ }
        }

        internal static void ApplyDataValidationForHeader(Excel.Worksheet ws, Excel.Range header, XqlSheet.Meta sm)
        {
            if (ws == null || header == null || sm == null) return;

            int colCount = 0, hdrRow = 1, hdrCol0 = 1;
            try { colCount = header.Columns.Count; } catch { colCount = 0; }
            try { hdrRow = header.Row; } catch { hdrRow = 1; }
            try { hdrCol0 = header.Column; } catch { hdrCol0 = 1; }
            if (colCount <= 0) return;

            var lo = XqlSheet.FindListObjectContaining(ws, header);

            int firstDataRow = hdrRow + 1;
            int lastDataRow = GetLastDataRowSafe(ws, header, firstDataRow, colCount);
            if (lastDataRow < firstDataRow) lastDataRow = firstDataRow;

            for (int i = 1; i <= colCount; i++)
            {
                using var h = SmartCom<Range>.Wrap((Excel.Range)header.Cells[1, i]);

                string? name = null;
                try { name = (h.Value?.Value2 as string)?.Trim(); } catch { name = null; }
                if (string.IsNullOrEmpty(name))
                {
                    int absCol = hdrCol0 + i - 1;
                    name = XqlCommon.ColumnIndexToLetter(absCol);
                }

                bool isIdCol = !string.IsNullOrWhiteSpace(sm.KeyColumn) &&
                               string.Equals(name, sm.KeyColumn, StringComparison.OrdinalIgnoreCase);

                try { if (h.Value != null) h.Value.Locked = isIdCol; } catch { }

                Excel.Range? target = null;
                int absColForThis = hdrCol0 + i - 1;

                if (lo?.DataBodyRange != null)
                {
                    try
                    {
                        var body = lo.DataBodyRange;
                        if (body != null && body.Rows.Count > 0)
                        {
                            int rel = absColForThis - body.Column + 1;
                            if (rel >= 1 && rel <= body.Columns.Count)
                                target = (Excel.Range)body.Columns[rel];
                        }
                    }
                    catch { target = null; }
                }

                if (target == null)
                {
                    try
                    {
                        var tl = (Excel.Range)ws.Cells[firstDataRow, absColForThis];
                        var br = (Excel.Range)ws.Cells[Math.Max(firstDataRow, lastDataRow), absColForThis];
                        target = ws.Range[tl, br];
                    }
                    catch { target = null; }
                }

                if (target != null && target.Rows.Count > 0)
                {
                    if (isIdCol)
                    {
                        try { LockIdColumn(ws, target); } catch { }
                        try { ApplyIdBlockedValidation(target); } catch { }
                    }
                    else
                    {
                        try { target.Validation.Delete(); } catch { }
                        try { target.Locked = false; } catch { }

                        if (!string.IsNullOrEmpty(name) && sm.Columns.TryGetValue(name!, out var ct))
                        {
                            try { ApplyValidationForKind(target, ct.Kind); } catch { }
                        }
                    }
                }
            }

            ApplyProtectionForHeaderAndIdOnly(ws, header, sm);

            static int GetLastDataRowSafe(Excel.Worksheet w, Excel.Range hdr, int firstRow, int headerColCount)
            {
                try
                {
                    var loX = XqlSheet.FindListObjectContaining(w, hdr);
                    var body = loX?.DataBodyRange;
                    if (body != null && body.Rows.Count > 0)
                        return body.Row + body.Rows.Count - 1;
                }
                catch { /* ignore */ }

                int last = firstRow - 1;
                for (int c = 0; c < Math.Max(1, headerColCount); c++)
                {
                    try
                    {
                        int absCol = hdr.Column + c;
                        using var lastCell = SmartCom<Excel.Range>.Acquire(() => (Excel.Range)w.Cells[w.Rows.Count, absCol]);
                        if (lastCell?.Value == null) continue;

                        using var hit = SmartCom<Excel.Range>.Acquire(() => (Excel.Range)lastCell.Value.End[Excel.XlDirection.xlUp]);
                        if (hit?.Value == null) continue;

                        int candidate = hit.Value.Row;
                        if (candidate >= firstRow && candidate > last) last = candidate;
                    }
                    catch { /* per-col ignore */ }
                }

                if (last < firstRow)
                {
                    try
                    {
                        using var used = SmartCom<Excel.Range>.Wrap(w.UsedRange);
                        try { _ = used?.Value?.Address[true, true, Excel.XlReferenceStyle.xlA1, false]; } catch { }
                        int usedLast = (used?.Value?.Row ?? 1) + (used?.Value?.Rows.Count ?? 1) - 1;
                        last = Math.Max(last, usedLast);
                    }
                    catch { }
                }
                return Math.Max(last, firstRow);
            }
        }

        // ─────────────────────────────────────────────────────────────────
        public static void MarkTouchedCell(Excel.Range rg)
        {
            if (rg == null) return;
            try
            {
                var interior = rg.Interior;
                interior.Pattern = Excel.XlPattern.xlPatternSolid;
                interior.Color = 0x00CCFFCC;
            }
            catch { /* ignore */ }
        }

        public static void MarkInvalidCell(Excel.Range rg)
        {
            if (rg == null) return;
            try
            {
                var interior = rg.Interior;
                interior.Pattern = Excel.XlPattern.xlPatternSolid;
                interior.Color = 0x00CCCCFF;
            }
            catch { /* ignore */ }
        }

        public static void TryClearInvalidMark(Excel.Range rg) => TryClearColor(rg, 0x00CCCCFF);
        public static void TryClearTouchedMark(Excel.Range rg) => TryClearColor(rg, 0x00CCFFCC);
        private static void TryClearColor(Excel.Range rg, int colorBgr)
        {
            if (rg == null) return;
            try
            {
                var it = rg.Interior;
                int cur = Convert.ToInt32(it.Color);
                if (cur == colorBgr)
                    it.ColorIndex = Excel.XlColorIndex.xlColorIndexNone;
            }
            catch { /* ignore */ }
        }

        public static void RecoverSummaryBegin()
        {
            lock (_sumLock)
            {
                _sumTables = new HashSet<string>(StringComparer.Ordinal);
                _sumAffected = _sumConflicts = _sumErrors = _sumBatches = 0;
                _sumStartTicks = System.Diagnostics.Stopwatch.GetTimestamp();
            }
        }

        public static void RecoverSummaryPush(string table, int affected, int conflicts, int errors)
        {
            lock (_sumLock)
            {
                _sumTables.Add(table ?? "");
                _sumAffected += Math.Max(0, affected);
                _sumConflicts += Math.Max(0, conflicts);
                _sumErrors += Math.Max(0, errors);
                _sumBatches++;
            }
        }

        public static void RecoverSummaryShow(string? title = "Recover Summary")
        {
            _ = XqlCommon.OnExcelThreadAsync(() =>
            {
                var app = (Excel.Application)ExcelDnaUtil.Application;
                using var wbW = SmartCom<Workbook>.Wrap(app.ActiveWorkbook);
                if (wbW.Value == null) return (object?)null;

                using var wsW = SmartCom<Worksheet>.Wrap(FindOrCreateSheet(wbW.Value, "_XQL_Summary"));
                if (wsW.Value == null) return (object?)null;

                try
                {
                    wsW.Value.Cells.ClearContents();
                    wsW.Value.Cells.ClearFormats();

                    int tables = _sumTables.Count;
                    double elapsedMs = TicksToMs(System.Diagnostics.Stopwatch.GetTimestamp() - _sumStartTicks);

                    Put(wsW.Value, 1, 1, title!, bold: true, size: 16);
                    Put(wsW.Value, 3, 1, "Tables"); Put(wsW.Value, 3, 2, tables.ToString());
                    Put(wsW.Value, 4, 1, "Batches"); Put(wsW.Value, 4, 2, _sumBatches.ToString());
                    Put(wsW.Value, 5, 1, "Affected Rows"); Put(wsW.Value, 5, 2, _sumAffected.ToString());
                    Put(wsW.Value, 6, 1, "Conflicts"); Put(wsW.Value, 6, 2, _sumConflicts.ToString());
                    Put(wsW.Value, 7, 1, "Errors"); Put(wsW.Value, 7, 2, _sumErrors.ToString());
                    Put(wsW.Value, 8, 1, "Elapsed (ms)"); Put(wsW.Value, 8, 2, elapsedMs.ToString("0"));

                    using var box = SmartCom<Range>.Wrap(wsW.Value.Range[wsW.Value.Cells[1, 1], wsW.Value.Cells[9, 3]]);
                    try
                    {
                        var interior = box.Value!.Interior;
                        interior.Pattern = Excel.XlPattern.xlPatternSolid;
                        interior.Color = 0x00F0F0F0;
                        box.Value!.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    }
                    catch { }

                    ((Excel.Range)wsW.Value.Columns["A:C"]).AutoFit();
                }
                catch { }

                return (object?)null;

                static void Put(Excel.Worksheet w, int r0, int c0, string text, bool bold = false, int? size = null)
                {
                    using var cell = SmartCom<Range>.Wrap((Excel.Range)w.Cells[r0, c0]);
                    if (cell.Value == null) return;
                    cell.Value.Value2 = text;
                    if (bold) cell.Value.Font.Bold = true;
                    if (size.HasValue) cell.Value.Font.Size = size.Value;
                }

                static double TicksToMs(long ticks)
                {
                    double freq = System.Diagnostics.Stopwatch.Frequency;
                    return ticks * 1000.0 / freq;
                }
            });
        }

        public static void AppendConflicts(IEnumerable<object>? conflicts)
        {
            if (conflicts == null) return;
            var items = conflicts.ToList();
            if (items.Count == 0) return;

            _ = XqlCommon.OnExcelThreadAsync(() =>
            {
                var app = (Excel.Application)ExcelDnaUtil.Application;
                using var wbW = SmartCom<Workbook>.Wrap(app.ActiveWorkbook);
                if (wbW.Value == null) return (object?)null;

                using var wsW = SmartCom<Worksheet>.Wrap(FindOrCreateSheet(wbW.Value, "_XQL_Conflicts"));
                if (wsW.Value == null) return (object?)null;

                try
                {
                    XqlSheet.TryGetUsedBounds(wsW.Value, out var fr, out var fc, out var lr, out var lc);
                    bool needHeader = (lr <= 1) || (((Excel.Range)wsW.Value.Cells[1, 1]).Value2 == null);

                    if (needHeader)
                    {
                        string[] headers = { "Timestamp", "Table", "RowKey", "Column", "Local", "Server", "Type", "Message", "Sheet", "Address" };
                        for (int i = 0; i < headers.Length; i++)
                            ((Excel.Range)wsW.Value.Cells[1, i + 1]).Value2 = headers[i];
                        using var hdr = SmartCom<Range>.Wrap(wsW.Value.Range[wsW.Value.Cells[1, 1], wsW.Value.Cells[1, headers.Length]]);
                        try { wsW.Value.ListObjects.Add(Excel.XlListObjectSourceType.xlSrcRange, hdr.Value, Type.Missing, Excel.XlYesNoGuess.xlYes); } catch { }
                        lr = Math.Max(lr, 1);
                    }

                    XqlSheet.TryGetUsedBounds(wsW.Value, out fr, out fc, out lr, out lc);
                    int last = lr;

                    foreach (var cf in items)
                    {
                        int next = Math.Max(2, last + 1);
                        using var row = SmartCom<Range>.Wrap(wsW.Value.Range[wsW.Value.Cells[next, 1], wsW.Value.Cells[next, 10]]);

                        string ts = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
                        string tbl = Prop(cf, "Table");
                        string rk = Prop(cf, "RowKey");
                        string col = Prop(cf, "Column");
                        string loc = Prop(cf, "Local") is string ? Prop(cf, "Local") : ToStr(PropObj(cf, "Local"));
                        string srv = Prop(cf, "Server") is string ? Prop(cf, "Server") : ToStr(PropObj(cf, "Server"));
                        string typ = Prop(cf, "Type");
                        string msg = Prop(cf, "Message");
                        string sh = Prop(cf, "Sheet");
                        string addr = Prop(cf, "Address");

                        ((Excel.Range)row.Value!.Cells[1, 1]).Value2 = ts;
                        ((Excel.Range)row.Value!.Cells[1, 2]).Value2 = tbl;
                        ((Excel.Range)row.Value!.Cells[1, 3]).Value2 = rk;
                        ((Excel.Range)row.Value!.Cells[1, 4]).Value2 = col;
                        ((Excel.Range)row.Value!.Cells[1, 5]).Value2 = loc;
                        ((Excel.Range)row.Value!.Cells[1, 6]).Value2 = srv;
                        ((Excel.Range)row.Value!.Cells[1, 7]).Value2 = typ;
                        ((Excel.Range)row.Value!.Cells[1, 8]).Value2 = msg;
                        ((Excel.Range)row.Value!.Cells[1, 9]).Value2 = sh;
                        ((Excel.Range)row.Value!.Cells[1, 10]).Value2 = addr;

                        try
                        {
                            var interior = row.Value!.Interior;
                            interior.Pattern = Excel.XlPattern.xlPatternSolid;
                            interior.Color = 0x00CCCCFF;
                        }
                        catch { }

                        if (!string.IsNullOrWhiteSpace(sh) && !string.IsNullOrWhiteSpace(addr))
                        {
                            try
                            {
                                string subAddr = $"'{sh.Replace("'", "''")}'!{addr}";
                                wsW.Value.Hyperlinks.Add(Anchor: row.Value!.Cells[1, 10], Address: "", SubAddress: subAddr, TextToDisplay: addr);
                            }
                            catch { }
                        }

                        last = next;
                    }
                }
                catch { }

                return (object?)null;

                static string Prop(object o, string name)
                    => Convert.ToString(PropObj(o, name), CultureInfo.InvariantCulture) ?? "";
                static object? PropObj(object o, string name)
                    => o.GetType().GetProperty(name, BindingFlags.Public | BindingFlags.Instance | BindingFlags.IgnoreCase)?.GetValue(o);
                static string ToStr(object? v)
                    => Convert.ToString(v, CultureInfo.InvariantCulture) ?? "";
            });
        }

        // ───────────────────────── 통합 엔트리: 헤더(plan) + 행 패치(patches)
        public static void ApplyPlanAndPatches(IReadOnlyDictionary<string, List<string>>? plan, IReadOnlyList<RowPatch>? patches)
        {
            if ((plan == null || plan.Count == 0) && (patches == null || patches.Count == 0))
                return;

            var tables = new HashSet<string>(StringComparer.Ordinal);
            if (plan is { Count: > 0 })
                foreach (var t in plan.Keys.Where(s => !string.IsNullOrWhiteSpace(s))) tables.Add(t);
            if (patches is { Count: > 0 })
                foreach (var t in patches.Select(p => p.Table).Where(s => !string.IsNullOrWhiteSpace(s))) tables.Add(t);

            var serverColsByTable = new Dictionary<string, List<ColumnInfo>>(StringComparer.Ordinal);
            try
            {
                if (XqlAddIn.Backend is IXqlBackend be)
                {
                    foreach (var tbl in tables)
                    {
                        try { serverColsByTable[tbl] = be.GetTableColumns(tbl).GetAwaiter().GetResult(); }
                        catch { }
                    }
                }
            }
            catch { }

            _ = XqlCommon.OnExcelThreadAsync(() =>
            {
                var app = (Excel.Application)ExcelDnaUtil.Application;
                if (app == null) return (object?)null;
                using var _batch = new XqlCommon.ExcelBatchScope(app);

                // 1) plan: 원하는 컬럼 스펙으로 각 테이블 헤더를 정확히 정렬·보장 (파이프라인 재사용)
                if (plan is { Count: > 0 })
                {
                    foreach (var (table, cols0) in plan)
                    {
                        if (string.IsNullOrWhiteSpace(table)) continue;
                        var cols = cols0?.Where(s => !string.IsNullOrWhiteSpace(s)).ToList() ?? new List<string>();
                        if (cols.Count == 0) continue;

                        EnsureHeaderForTable(app, table, cols); // ← 내부에서 ExecuteHeaderPipeline(HeaderOp.Refresh, spec) 사용
                    }
                }

                // 2) 서버 스키마 → 메타 반영 + 헤더 UI/Validation 재적용
                foreach (var tbl in tables)
                {
                    try
                    {
                        using var wsW = SmartCom<Worksheet>.Wrap(XqlSheet.FindWorksheetByTable(app, tbl, out var smeta));
                        if (wsW.Value == null || smeta == null) continue;

                        if (serverColsByTable.TryGetValue(tbl, out var infos) && infos != null)
                            ApplyServerColumnsToMeta(smeta, infos);
                        else if (string.IsNullOrWhiteSpace(smeta.KeyColumn))
                            smeta.KeyColumn = "id";

                        var lo = XqlSheet.FindListObjectByTable(wsW.Value, tbl);
                        using var headerW = SmartCom<Range>.Wrap(lo?.HeaderRowRange ?? XqlSheet.GetHeaderRange(wsW.Value));
                        if (headerW.Value != null)
                        {
                            ApplyHeaderUi(wsW.Value, headerW.Value, smeta, withValidation: true);
                        }
                    }
                    catch { }
                }

                // 3) 패치 적용 + 지문(UID) 기록
                if (patches is { Count: > 0 })
                {
                    InternalApplyCore(app, patches);
                    AppendFingerprintsForPatches(app, patches);
                }

                return (object?)null;
            });
        }

        /// <summary>
        /// 헤더 영역에서 keyName 열을 첫번째로 보정하고, 새 헤더 Range와 컬럼명 목록을 반환.
        /// </summary>
        private static (Excel.Range newHeader, List<string> names) EnsureIdFirstAndRebuildHeader(
            Excel.Worksheet ws, Excel.Range header, string keyName)
        {
            if (ws == null || header == null) throw new ArgumentNullException(nameof(header));

            using var headerW = SmartCom<Excel.Range>.Wrap(header);
            int row0 = headerW.Value!.Row, col0 = headerW.Value.Column, cols0 = headerW.Value.Columns.Count;

            var names = BuildHeaderNames(headerW.Value);

            int idIdx1 = -1;
            for (int i = 0; i < names.Count; i++)
                if (string.Equals(names[i], keyName, StringComparison.OrdinalIgnoreCase)) { idIdx1 = i + 1; break; }

            if (idIdx1 < 0)
            {
                using var idCell = SmartCom<Excel.Range>.Wrap((Excel.Range)ws.Cells[row0, col0]);
                try { if (idCell.Value != null) idCell.Value.Value2 = keyName; } catch { }
            }
            else if (idIdx1 != 1)
            {
                int idAbsCol = col0 + (idIdx1 - 1);
                using var srcCol = SmartCom<Excel.Range>.Wrap((Excel.Range)ws.Columns[idAbsCol]);
                using var dest = SmartCom<Excel.Range>.Wrap((Excel.Range)ws.Cells[row0, col0]);
                try
                {
                    srcCol.Value?.Copy(dest.Value);
                    srcCol.Value?.Clear();
                }
                catch { }
            }

            using var used = SmartCom<Excel.Range>.Wrap(ws.UsedRange as Excel.Range);
            var start = SmartCom<Excel.Range>.Wrap((Excel.Range)ws.Cells[row0, col0]);
            var end = SmartCom<Excel.Range>.Wrap((Excel.Range)ws.Cells[row0, col0 + Math.Max(0, cols0 - 1)]);
            using var newHeaderW = SmartCom<Excel.Range>.Wrap(ws.Range[start.Value, end.Value]);

            var newNames = BuildHeaderNames(newHeaderW.Value!);
            return (newHeaderW.Detach()!, newNames);
        }

        private static void InternalApplyCore(Excel.Application app, IReadOnlyList<RowPatch> patches)
        {
            foreach (var grp in patches.GroupBy(p => p.Table, StringComparer.Ordinal))
            {
                using var wsW = SmartCom<Worksheet>.Wrap(XqlSheet.FindWorksheetByTable(app, grp.Key, out var smeta));
                if (wsW.Value == null || smeta == null) continue;

                var lo = XqlSheet.FindListObjectByTable(wsW.Value, grp.Key);
                using var headerW = SmartCom<Range>.Wrap(lo?.HeaderRowRange ?? XqlSheet.GetHeaderRange(wsW.Value));
                if (headerW.Value == null) continue;

                var headers = XqlSheet.ComputeHeaderNames(headerW.Value);

                var serverCols = new HashSet<string>(StringComparer.Ordinal);
                foreach (var p in grp)
                {
                    if (p.Deleted || p.Cells == null) continue;
                    foreach (var k in p.Cells.Keys)
                        if (!string.IsNullOrWhiteSpace(k))
                            serverCols.Add(k);
                }
                var keyName = string.IsNullOrWhiteSpace(smeta.KeyColumn) ? "id" : smeta.KeyColumn!;
                serverCols.Add(keyName);

                bool needCreateHeader =
                    headers.Count == 0 ||
                    XqlSheet.IsFallbackLetterHeader(headerW.Value) ||
                    !headers.Any(h => serverCols.Contains(h));

                Excel.Range headerFinal = headerW.Value;
                if (needCreateHeader && serverCols.Count > 0)
                {
                    var ordered = new List<string>(serverCols.Count);
                    if (serverCols.Contains(keyName)) ordered.Add(keyName);
                    ordered.AddRange(serverCols.Where(c => !string.Equals(c, keyName, StringComparison.Ordinal))
                                               .OrderBy(c => c, StringComparer.Ordinal));

                    headerFinal = UpdateHeaderToColumns(wsW.Value, headerW.Value, smeta, grp.Key, ordered);
                    headers = ordered;
                }
                if (headers.Count == 0) continue;

                try { XqlAddIn.Sheet!.EnsureColumns(wsW.Value.Name, serverCols.ToArray()); } catch { }

                int keyIdx1 = XqlSheet.FindKeyColumnIndex(headers, smeta.KeyColumn);
                int keyAbsCol = headerFinal.Column + keyIdx1 - 1;
                int firstDataRow = headerFinal.Row + 1;

                foreach (var patch in grp)
                {
                    try
                    {
                        int? row = XqlSheet.FindRowByKey(wsW.Value, firstDataRow, keyAbsCol, patch.RowKey);
                        if (patch.Deleted)
                        {
                            if (row.HasValue) SafeDeleteRow(wsW.Value, row.Value);
                            continue;
                        }
                        if (!row.HasValue) row = AppendNewRow(wsW.Value, firstDataRow, lo);

                        ApplyCells(wsW.Value, row!.Value, headerFinal, headers, smeta, patch.Cells);
                    }
                    catch { /* per-row 안전 */ }

                    try { ApplyHeaderUi(wsW.Value, headerFinal, smeta, withValidation: true); } catch { }
                }
            }
        }

        private static Excel.Range UpdateHeaderToColumns(Excel.Worksheet ws, Excel.Range oldHeader, XqlSheet.Meta smeta, string tableName, IList<string> columns)
        {
            string keyName = string.IsNullOrWhiteSpace(smeta.KeyColumn) ? "id" : smeta.KeyColumn!;
            var cols = new List<string>(columns.Count + 1);
            if (!columns.Any(c => c.Equals(keyName, StringComparison.OrdinalIgnoreCase)))
                cols.Add(keyName);
            cols.AddRange(columns.Where(c => !c.Equals(keyName, StringComparison.OrdinalIgnoreCase)));

            using var start = SmartCom<Range>.Wrap((Excel.Range)ws.Cells[oldHeader.Row, oldHeader.Column]);
            using var end = SmartCom<Range>.Wrap((Excel.Range)ws.Cells[oldHeader.Row, oldHeader.Column + cols.Count - 1]);
            using var newHeader = SmartCom<Range>.Wrap(ws.Range[start.Value, end.Value]);

            var arr = new object[1, cols.Count];
            for (int i = 0; i < cols.Count; i++) arr[0, i] = cols[i] ?? "";
            newHeader.Value!.Value2 = arr;

            smeta.KeyColumn = keyName;
            XqlAddIn.Sheet!.EnsureColumns(ws.Name, cols);
            XqlSheet.SetHeaderMarker(ws, newHeader.Value!);
            ApplyHeaderUi(ws, newHeader.Value!, smeta, withValidation: true);
            InvalidateHeaderCache(ws.Name);
            RegisterTableSheet(tableName, ws.Name);

            return newHeader.Detach()!;
        }

        // ✅ 시트 보장 + 파이프라인 재사용(스펙 적용 Refresh)
        private static void EnsureHeaderForTable(Excel.Application app, string table, List<string> columns)
        {
            using var wsW = SmartCom<Worksheet>.Wrap(XqlSheet.FindWorksheet(app, table) ?? CreateSheet(app, table));
            if (wsW.Value == null) return;

            var spec = new HeaderSpec(columns);
            ExecuteHeaderPipeline(app, wsW.Value!, HeaderOp.Refresh, spec); // 헤더/툴팁/유효성/보호/마커까지 일괄
        }

        // 시트 생성 유틸
        private static Excel.Worksheet CreateSheet(Excel.Application app, string name)
        {
            var sheets = app.Worksheets;
            var last = (Excel.Worksheet)sheets[sheets.Count];
            var newWs = (Excel.Worksheet)sheets.Add(After: last);
            try { newWs.Name = name; } catch { }
            return newWs;
        }

        // ✅ 서버 스키마 → 메타 반영 헬퍼
        private static void ApplyServerColumnsToMeta(XqlSheet.Meta smeta, IEnumerable<ColumnInfo> infos)
        {
            if (smeta == null || infos == null) return;

            string? pkName = infos.FirstOrDefault(c => c.pk)?.name;
            if (!string.IsNullOrWhiteSpace(pkName))
                smeta.KeyColumn = pkName!;
            if (string.IsNullOrWhiteSpace(smeta.KeyColumn))
                smeta.KeyColumn = "id";

            foreach (var info in infos)
            {
                var n = (info.name ?? "").Trim();
                if (n.Length == 0) continue;

                var t = (info.type ?? "").Trim().ToUpperInvariant();
                var ct = smeta.Columns.TryGetValue(n, out var cur) && cur != null ? cur : new XqlSheet.ColumnType();
                ct.Kind = t switch
                {
                    "INT" or "INTEGER" => XqlSheet.ColumnKind.Int,
                    "REAL" or "FLOAT" or "DOUBLE" or "NUMERIC" => XqlSheet.ColumnKind.Real,
                    "BOOL" or "BOOLEAN" => XqlSheet.ColumnKind.Bool,
                    "DATE" or "DATETIME" or "TIMESTAMP" => XqlSheet.ColumnKind.Date,
                    "JSON" => XqlSheet.ColumnKind.Json,
                    _ => XqlSheet.ColumnKind.Text
                };
                ct.Nullable = !info.notnull;

                smeta.SetColumn(n, ct);
            }
        }

        private static void AppendFingerprintsForPatches(Excel.Application app, IReadOnlyList<RowPatch> patches)
        {
            try
            {
                var wb = app.ActiveWorkbook;
                var items = new List<(string table, string rowKey, string colUid, string fp)>(Math.Max(64, patches.Count));

                foreach (var grp in patches.GroupBy(p => p.Table, StringComparer.Ordinal))
                {
                    using var wsW = SmartCom<Worksheet>.Wrap(XqlSheet.FindWorksheetByTable(app, grp.Key, out var smeta));
                    if (wsW.Value == null) continue;

                    var lo = XqlSheet.FindListObjectByTable(wsW.Value, grp.Key);
                    using var headerW = SmartCom<Range>.Wrap(lo?.HeaderRowRange ?? XqlSheet.GetHeaderRange(wsW.Value));
                    if (headerW.Value == null) continue;

                    var uidMap = GetUidMapCached(wsW.Value, headerW.Value);

                    foreach (var rp in grp)
                    {
                        if (rp.Deleted || rp.Cells == null || rp.Cells.Count == 0) continue;
                        foreach (var kv in rp.Cells)
                        {
                            var col = kv.Key;
                            if (!uidMap.TryGetValue(col, out var uid) || string.IsNullOrEmpty(uid)) continue;
                            var fp = XqlCommon.Fingerprint(kv.Value);
                            items.Add((grp.Key, Convert.ToString(rp.RowKey) ?? "", uid, fp));
                        }
                    }
                }

                if (items.Count > 0) XqlSheet.ShadowAppendFingerprints(wb, items);
            }
            catch { /* 무음 */ }
        }

        // 문자열/맵 캐시 (COM X)
        internal static bool TryGetCachedSheetForTable(string table, out string sheetName)
            => _tableToSheet.TryGetValue(table, out sheetName!);

        internal static void CacheTableToSheet(string table, string sheetName)
        {
            if (!string.IsNullOrWhiteSpace(table) && !string.IsNullOrWhiteSpace(sheetName))
                _tableToSheet[table] = sheetName;
        }

        private static Dictionary<string, string> GetUidMapCached(Excel.Worksheet ws, Excel.Range header)
        {
            string addr;
            try { addr = header.Address[true, true, Excel.XlReferenceStyle.xlA1, Type.Missing, Type.Missing]; }
            catch { addr = $"{ws.Name}:{header.Row}:{header.Column}:{header.Columns.Count}"; }

            if (_hdrCache.TryGetValue(ws.Name, out var e) && e.addr == addr) return e.map;

            var map = XqlSheet.BuildColUidMap(ws, header);
            _hdrCache[ws.Name] = (addr, map);
            return map;
        }

        public static void InvalidateHeaderCache(string sheetName)
        {
            _hdrCache.TryRemove(sheetName, out _);
        }

        public static void RegisterTableSheet(string table, string sheetName)
        {
            if (!string.IsNullOrWhiteSpace(table) && !string.IsNullOrWhiteSpace(sheetName))
                _tableToSheet[table] = sheetName;
        }

        private static int AppendNewRow(Excel.Worksheet ws, int firstDataRow, Excel.ListObject? lo)
        {
            if (lo != null)
            {
                try
                {
                    var lr = lo.ListRows.Add();
                    var body = lo.DataBodyRange;
                    if (body != null)
                    {
                        int row = body.Row + body.Rows.Count - 1;
                        return row;
                    }
                }
                catch { /* 폴백 */ }
            }

            XqlSheet.TryGetUsedBounds(ws, out var fr, out var fc, out var lr2, out var lc2);
            int last = lr2;
            return Math.Max(firstDataRow, last + 1);
        }

        private static void ApplyCells(Excel.Worksheet ws, int row, Excel.Range header,
                                       List<string> headers, XqlSheet.Meta meta, Dictionary<string, object?> cells)
        {
            for (int c = 0; c < headers.Count; c++)
            {
                var colName = headers[c];
                if (string.IsNullOrWhiteSpace(colName)) continue;

                if (!meta.Columns.ContainsKey(colName))
                {
                    try
                    {
                        meta.SetColumn(colName, new XqlSheet.ColumnType
                        {
                            Kind = XqlSheet.ColumnKind.Text,
                            Nullable = true
                        });
                    }
                    catch { }
                }

                if (!cells.TryGetValue(colName, out var val)) continue;

                using var rg = SmartCom<Range>.Wrap((Excel.Range)ws.Cells[row, header.Column + c]);
                if (rg.Value == null) continue;

                try
                {
                    if (val == null) { rg.Value.Value2 = null; MarkTouchedCell(rg.Value); continue; }

                    switch (val)
                    {
                        case bool b: rg.Value.Value2 = b; break;
                        case long l: rg.Value.Value2 = (double)l; break;
                        case int i: rg.Value.Value2 = (double)i; break;
                        case double d: rg.Value.Value2 = d; break;
                        case float f: rg.Value.Value2 = (double)f; break;
                        case decimal m: rg.Value.Value2 = (double)m; break;
                        case DateTime dt: rg.Value.Value2 = dt.ToOADate(); break;
                        default:
                            rg.Value.Value2 = Convert.ToString(val, System.Globalization.CultureInfo.InvariantCulture);
                            break;
                    }
                    MarkTouchedCell(rg.Value);
                }
                catch (Exception ex)
                {
                    XqlLog.Error($"패치 적용 실패: {ex.Message}", ws.Name, rg.Value?.Address[false, false] ?? "");
                }
            }
        }

        private static void SafeDeleteRow(Excel.Worksheet ws, int row)
        {
            try { ((Excel.Range)ws.Rows[row]).Delete(); }
            catch { }
        }

        private static Excel.Worksheet FindOrCreateSheet(Excel.Workbook wb, string name)
        {
            foreach (Excel.Worksheet s in wb.Worksheets)
            {
                if (string.Equals(s.Name, name, StringComparison.Ordinal))
                    return s;
            }
            var created = (Excel.Worksheet)wb.Worksheets.Add();
            created.Name = name;
            created.Move(After: wb.Worksheets[wb.Worksheets.Count]);
            return created;
        }

        private static void ApplyHeaderOutlineBorder(Excel.Range header)
        {
            using var bs = SmartCom<Borders>.Wrap(header.Borders);
            if (bs.Value == null) return;

            var idxs = new[]
            {
                Excel.XlBordersIndex.xlEdgeLeft,
                Excel.XlBordersIndex.xlEdgeTop,
                Excel.XlBordersIndex.xlEdgeRight,
                Excel.XlBordersIndex.xlEdgeBottom
            };
            foreach (var idx in idxs)
            {
                var b = bs.Value[idx];
                try
                {
                    b.LineStyle = Excel.XlLineStyle.xlContinuous;
                    b.Weight = Excel.XlBorderWeight.xlThin;
                }
                catch { }
            }
        }

        private static void ClearHeaderUi(Excel.Worksheet ws, Excel.Range? header, bool removeMarker = false)
        {
            using var hdrW = SmartCom<Range>.Wrap(header ?? XqlSheet.GetHeaderRange(ws));
            if (hdrW.Value == null) return;

            foreach (Excel.Range cell in hdrW.Value.Cells)
            {
                try
                {
                    try { cell.ClearComments(); } catch { try { cell.Comment?.Delete(); } catch { } }
                }
                catch { /* ignore */ }
            }

            try
            {
                var bs = hdrW.Value.Borders;
                var idxs = new[]
                {
                    Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBordersIndex.xlEdgeTop,
                    Excel.XlBordersIndex.xlEdgeRight, Excel.XlBordersIndex.xlEdgeBottom,
                    Excel.XlBordersIndex.xlInsideHorizontal, Excel.XlBordersIndex.xlInsideVertical
                };
                foreach (var idx in idxs)
                {
                    try { bs[idx].LineStyle = Excel.XlLineStyle.xlLineStyleNone; } catch { }
                }
            }
            catch { }

            try
            {
                foreach (Excel.Range h in hdrW.Value.Cells)
                {
                    using var first = SmartCom<Range>.Wrap((Excel.Range)h.Offset[1, 0]);
                    using var last = SmartCom<Range>.Wrap((Excel.Range)ws.Cells[ws.Rows.Count, h.Column]);
                    using var col = SmartCom<Range>.Wrap(ws.Range[first.Value, last.Value]);
                    try { col.Value?.Validation.Delete(); } catch { }

                    if (col.Value != null)
                    {
                        foreach (Excel.Range c in col.Value.Cells)
                        {
                            try { TryClearInvalidMark(c); TryClearTouchedMark(c); }
                            catch { }
                        }
                    }
                }
            }
            catch { }

            if (removeMarker) XqlSheet.ClearHeaderMarker(ws);
        }

        private static Excel.Range ColBelowToEnd(Excel.Worksheet ws, Excel.Range h)
        {
            using var first = SmartCom<Range>.Wrap((Excel.Range)h.Offset[1, 0]);
            using var last = SmartCom<Range>.Wrap((Excel.Range)ws.Cells[ws.Rows.Count, h.Column]);
            return SmartCom<Range>.Wrap(ws.Range[first.Value, last.Value]).Detach()!;
        }

        private static void ApplyValidationForKind(Excel.Range rng, XqlSheet.ColumnKind kind)
        {
            if (rng == null) return;

            try
            {
                var areas = rng.Areas;
                if (areas != null && areas.Count > 1)
                {
                    foreach (Excel.Range a in areas)
                    {
                        try { ApplyValidationForKind(a, kind); }
                        finally { ReleaseCom(a); }
                    }
                    return;
                }
            }
            catch { /* ignore */ }

            using var wsW = SmartCom<Worksheet>.Wrap(rng.Worksheet);
            var ws = wsW.Value;
            bool needReprotect = false;
            try
            {
                if (ws != null && ws.ProtectContents)
                {
                    needReprotect = true;
                    ws.Unprotect(Type.Missing);
                }
            }
            catch { /* ignore */ }

            try
            {
                try
                {
                    if ((long)rng.CountLarge == 0) return;
                }
                catch { /* ignore */ }

                try { rng.Validation?.Delete(); } catch { }

                using var vW = XqlCommon.SmartCom<Excel.Validation>.Wrap(rng.Validation);
                if (vW.Value == null) return;

                bool added = false;

                switch (kind)
                {
                    case XqlSheet.ColumnKind.Int:
                        vW.Value.Add(
                            Excel.XlDVType.xlValidateWholeNumber,
                            Excel.XlDVAlertStyle.xlValidAlertStop,
                            Excel.XlFormatConditionOperator.xlBetween,
                            (double)int.MinValue, (double)int.MaxValue);
                        vW.Value.IgnoreBlank = true;
                        vW.Value.ErrorTitle = "정수만 허용";
                        vW.Value.ErrorMessage = "이 열은 정수만 입력할 수 있습니다.";
                        added = true;
                        break;

                    case XqlSheet.ColumnKind.Real:
                        vW.Value.Add(
                            Excel.XlDVType.xlValidateDecimal,
                            Excel.XlDVAlertStyle.xlValidAlertStop,
                            Excel.XlFormatConditionOperator.xlBetween,
                            -1e307, 1e307);
                        vW.Value.IgnoreBlank = true;
                        vW.Value.ErrorTitle = "실수만 허용";
                        vW.Value.ErrorMessage = "이 열은 실수/숫자만 입력할 수 있습니다.";
                        added = true;
                        break;

                    case XqlSheet.ColumnKind.Bool:
                        string listSep = ",";
                        try
                        {
                            var app = (Excel.Application)rng.Application;
                            if (app.International[Excel.XlApplicationInternational.xlListSeparator] is string s && s.Length > 0)
                                listSep = s;
                        }
                        catch { }
                        vW.Value.Add(
                            Excel.XlDVType.xlValidateList,
                            Excel.XlDVAlertStyle.xlValidAlertStop,
                            Type.Missing, $"TRUE{listSep}FALSE", Type.Missing);
                        vW.Value.IgnoreBlank = true;
                        vW.Value.ErrorTitle = "BOOL만 허용";
                        vW.Value.ErrorMessage = "이 열은 TRUE 또는 FALSE만 입력할 수 있습니다.";
                        added = true;
                        break;

                    case XqlSheet.ColumnKind.Date:
                        double dmin = new DateTime(1900, 1, 1).ToOADate();
                        double dmax = new DateTime(9999, 12, 31).ToOADate();
                        vW.Value.Add(
                            Excel.XlDVType.xlValidateDate,
                            Excel.XlDVAlertStyle.xlValidAlertStop,
                            Excel.XlFormatConditionOperator.xlBetween,
                            dmin, dmax);
                        vW.Value.IgnoreBlank = true;
                        vW.Value.ErrorTitle = "날짜만 허용";
                        vW.Value.ErrorMessage = "이 열은 날짜 형식만 입력할 수 있습니다.";
                        added = true;
                        break;

                    default:
                        added = false;
                        break;
                }

                if (added)
                {
                    try { vW.Value.ShowError = true; } catch { }
                    try { vW.Value.ShowInput = false; } catch { }
                }
            }
            finally
            {
                if (needReprotect && ws != null)
                    try
                    {
                        using var headerW = SmartCom<Range>.Wrap(GetHeaderOrFallback(ws));
                        var sm = XqlAddIn.Sheet?.GetOrCreateSheet(ws.Name);
                        if (headerW.Value != null && sm != null)
                            ApplyProtectionForHeaderAndIdOnly(ws, headerW.Value, sm);
                        else
                            EnsureSheetProtectedUiOnly(ws);
                    }
                    catch { }
            }
        }

        internal static Excel.Range? GetHeaderOrFallback(Excel.Worksheet ws)
        {
            if (XqlSheet.TryGetHeaderMarker(ws, out var hdr)) return hdr;
            using var selW = SmartCom<Range>.Wrap(GetSelection(ws));
            return ResolveHeader(ws, selW.Value, XqlAddIn.Sheet!) ?? XqlSheet.GetHeaderRange(ws);
        }

        private static void RebindMarkerIfMoved(Excel.Worksheet ws, Excel.Range candidate)
        {
            if (XqlSheet.TryGetHeaderMarker(ws, out var old))
            {
                using var _ = SmartCom<Range>.Wrap(old);
                if (!XqlSheet.IsSameRange(old, candidate)) XqlSheet.SetHeaderMarker(ws, candidate);
            }
            else
            {
                XqlSheet.SetHeaderMarker(ws, candidate);
            }
        }

        // 단일 진입점: 헤더 UI(툴팁/테두리/유효성/보호)
        internal static void ApplyHeaderUi(Excel.Worksheet ws, Excel.Range header, XqlSheet.Meta sm, bool withValidation)
        {
            if (ws == null || header == null || sm == null) return;

            var tips = BuildHeaderTooltips(sm, header);
            SetHeaderTooltips(header, tips);
            ApplyHeaderOutlineBorder(header);

            if (withValidation)
                ApplyDataValidationForHeader(ws, header, sm);
        }

        /// <summary>id 컬럼 잠금</summary>
        private static void LockIdColumn(Excel.Worksheet ws, Excel.Range colData)
        {
            try { colData.Locked = true; } catch { }
        }

        /// <summary>ID 차단 Validation(로케일 독립식, 보호 해제/원복 포함)</summary>
        internal static void ApplyIdBlockedValidation(Excel.Range rng)
        {
            if (rng == null) return;

            Excel.Worksheet? ws = null; bool reProtect = false;
            try { ws = rng.Worksheet; } catch { }

            try
            {
                if (ws != null && ws.ProtectContents)
                {
                    ws.Unprotect(Type.Missing);
                    reProtect = true;
                }
            }
            catch { /* ignore */ }

            void One(Excel.Range r)
            {
                if (r == null) return;
                try { if (r.MergeCells is bool m && m) return; } catch { }
                try { r.Validation?.Delete(); } catch { }
                try
                {
                    var v = r.Validation;
                    v!.Add(Excel.XlDVType.xlValidateCustom, Excel.XlDVAlertStyle.xlValidAlertStop, Type.Missing, "=1=0", Type.Missing);
                    v.ErrorTitle = "편집 차단";
                    v.ErrorMessage = "ID 열은 서버가 관리합니다.";
                    v.ShowError = true;
                    v.IgnoreBlank = true;
                    try { r.Locked = true; } catch { }
                }
                catch
                {
                    try { r.Validation?.Delete(); } catch { }
                    try { r.Locked = true; } catch { }
                }
            }

            try
            {
                if ((rng.Areas?.Count ?? 1) > 1)
                    foreach (Excel.Range a in rng.Areas!) One(a);
                else
                    One(rng);
            }
            finally
            {
                if (reProtect && ws != null)
                {
                    try
                    {
                        using var headerW = SmartCom<Excel.Range>.Wrap(GetHeaderOrFallback(ws));
                        var sm = XqlAddIn.Sheet?.GetOrCreateSheet(ws.Name);
                        if (headerW.Value != null && sm != null)
                            ApplyProtectionForHeaderAndIdOnly(ws, headerW.Value, sm);
                        else
                            EnsureSheetProtectedUiOnly(ws);
                    }
                    catch { /* ignore */ }
                }
            }
        }

        private static void EnsureSheetProtectedUiOnly(Excel.Worksheet ws)
        {
            try { ws.Unprotect(Type.Missing); } catch { }
            try
            {
                ws.Protect(
                    Password: Type.Missing,
                    DrawingObjects: false,
                    Contents: true,
                    Scenarios: false,
                    UserInterfaceOnly: true,
                    AllowFormattingCells: true,
                    AllowFormattingColumns: true,
                    AllowFiltering: true,
                    AllowSorting: true
                );
                try { ws.EnableSelection = Excel.XlEnableSelection.xlUnlockedCells; } catch { }
            }
            catch { /* 무음 */ }
        }

        private static void ApplyProtectionForHeaderAndIdOnly(Excel.Worksheet ws, Excel.Range header, XqlSheet.Meta sm)
        {
            if (ws == null || header == null || sm == null) return;

            try { ws.Unprotect(Type.Missing); } catch { }

            int colCount = 0;
            try { colCount = header.Columns.Count; } catch { colCount = 0; }
            if (colCount <= 0) { EnsureSheetProtectedUiOnly(ws); return; }

            for (int i = 1; i <= colCount; i++)
            {
                using var h = SmartCom<Range>.Wrap((Excel.Range)header.Cells[1, i]);
                string? name = null;
                try { name = (h.Value?.Value2 as string)?.Trim(); } catch { }

                bool isIdCol = !string.IsNullOrWhiteSpace(sm.KeyColumn) &&
                               string.Equals(name, sm.KeyColumn, StringComparison.OrdinalIgnoreCase);

                try { if (h.Value != null) h.Value.Locked = isIdCol; } catch { }

                using var body = SmartCom<Range>.Wrap(ColBelowToEnd(ws, h.Value!));
                if (body.Value == null) continue;

                try { body.Value.Locked = isIdCol; } catch { }
            }

            EnsureSheetProtectedUiOnly(ws);
        }

        // ───────────────────────── 미래 대비: 콜라보 락 반영 훅
        internal sealed class CollabLock
        {
            public string ResourceKey { get; set; } = "";
            public string Owner { get; set; } = "";
        }

        public static void ApplyCollabLocks(string sheetName, IEnumerable<CollabLock> locks, string? myOwner = null)
        {
            if (string.IsNullOrWhiteSpace(sheetName)) return;
            var items = locks?.ToList() ?? new();

            _ = XqlCommon.OnExcelThreadAsync(() =>
            {
                var app = (Excel.Application)ExcelDnaUtil.Application;
                using var wsW = SmartCom<Excel.Worksheet>.Wrap(XqlSheet.FindWorksheet(app, sheetName));
                if (wsW.Value == null) return (object?)null;

                using var headerW = SmartCom<Excel.Range>.Wrap(GetHeaderOrFallback(wsW.Value));
                var sm = XqlAddIn.Sheet?.GetOrCreateSheet(sheetName);
                if (headerW.Value == null || sm == null) return (object?)null;

                ApplyProtectionForHeaderAndIdOnly(wsW.Value, headerW.Value, sm);

                var foreignLocks = string.IsNullOrWhiteSpace(myOwner)
                    ? items
                    : items.Where(l => !string.Equals(l.Owner, myOwner, StringComparison.Ordinal)).ToList();

                if (foreignLocks.Count == 0) return (object?)null;

                try { wsW.Value.Unprotect(Type.Missing); } catch { }

                foreach (var lk in foreignLocks)
                {
                    try
                    {
                        if (!XqlSheet.TryParse(lk.ResourceKey, out var desc)) continue;
                        if (!XqlSheet.TryResolve(app, desc, out var target, out _, out _)) continue;
                        using var tW = SmartCom<Excel.Range>.Wrap(target);

                        if (tW.Value != null)
                        {
                            if (string.Equals(desc.Kind, "col", StringComparison.OrdinalIgnoreCase))
                            {
                                using var first = SmartCom<Excel.Range>.Wrap((Excel.Range)tW.Value.Offset[1, 0]);
                                using var last = SmartCom<Excel.Range>.Wrap((Excel.Range)wsW.Value.Cells[wsW.Value.Rows.Count, tW.Value.Column]);
                                using var body = SmartCom<Excel.Range>.Wrap(wsW.Value.Range[first.Value, last.Value]);
                                if (body.Value != null) body.Value.Locked = true;
                            }
                            else
                            {
                                tW.Value.Locked = true;
                            }
                        }
                    }
                    catch { /* ignore per lock */ }
                }

                EnsureSheetProtectedUiOnly(wsW.Value);
                return (object?)null;
            });
        }
    }
}
