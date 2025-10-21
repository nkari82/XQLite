// XqlSheetView.cs
using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
using System.Windows.Forms;
using XQLite.AddIn;
using static XQLite.AddIn.XqlSchemaForm;
using static XQLite.AddIn.XqlSheet;
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
        private static readonly object _sumLock = new();
        private static HashSet<string> _sumTables = new(StringComparer.Ordinal);
        private static int _sumAffected, _sumConflicts, _sumErrors, _sumBatches;
        private static long _sumStartTicks;
        private static readonly ConcurrentDictionary<string, string> _tableToSheet = new(StringComparer.Ordinal);
        private static readonly ConcurrentDictionary<string, (string addr, Dictionary<string, string> map)> _hdrCache = new(StringComparer.Ordinal);

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

                // 메타 동기화 + ID 1열 강제
                var names = BuildHeaderNames(candidate);
                var sm = sheet.GetOrCreateSheet(ws.Name);

                // 항상 KeyColumn = "id" 보장
                if (string.IsNullOrWhiteSpace(sm.KeyColumn) || !string.Equals(sm.KeyColumn, "id", StringComparison.OrdinalIgnoreCase))
                {
                    sm.KeyColumn = "id";
                }

                // 헤더에 id가 맨 앞에 오도록 열 자체를 정렬(필요 시 이동/추가)
                candidate = EnsureIdIsFirst(ws, candidate, sm, names, reorderData: true);
                // 최종 이름 재수집
                names = BuildHeaderNames(candidate);
                sheet.EnsureColumns(ws.Name, names);

                // UI/검증 한 번에
                ApplyHeaderUi(ws, candidate, sm, withValidation: true);

                // 마커 확정
                XqlSheet.SetHeaderMarker(ws, candidate);

                // Excel 내부 후처리 이후 유실 방지(지연 재적용·중복 큐잉 방지)
                EnqueueReapplyHeaderUi(ws.Name, withValidation: true);

                InvalidateHeaderCache(ws.Name);
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("InstallHeader failed: " + ex.Message, "XQLite");
                return false;
            }
            finally { XqlCommon.ReleaseCom(candidate, ws); }
        }

        // 메타에 있으면 메타 기반, 없으면 폴백
        private static string ColumnTooltipFor(XqlSheet.Meta sm, string colName)
        {
            try
            {
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

        // 헤더 범위를 돌며 i열(1-based) → 툴팁 텍스트를 만든다.
        internal static IReadOnlyDictionary<int, string> BuildHeaderTooltips(XqlSheet.Meta sm, Excel.Range header)
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

                    tips[i] = ColumnTooltipFor(sm, colName!);
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

                // 갱신 시에도 id를 1열로 유지(필요 시 열 이동)
                header = EnsureIdIsFirst(ws, header, sm, BuildHeaderNames(header), reorderData: true);

                ApplyHeaderUi(ws, header, sm, withValidation: true);

                EnqueueReapplyHeaderUi(ws.Name, withValidation: true);

                InvalidateHeaderCache(ws.Name);
            }
            catch (Exception ex) { MessageBox.Show("RefreshMetaHeader failed: " + ex.Message, "XQLite"); }
            finally { XqlCommon.ReleaseCom(header, ws); }
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
            finally { XqlCommon.ReleaseCom(sel, hdr, ws); }
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
            finally { XqlCommon.ReleaseCom(header, ws); }
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

        // 헤더명 추출은 공용 유틸 사용
        private static List<string> BuildHeaderNames(Excel.Range header)
            => XqlSheet.ComputeHeaderNames(header);

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

                        if (string.IsNullOrEmpty(text))
                        {
                            try { cell.Comment?.Delete(); } catch { }
                            continue;
                        }

                        cmt = cell.Comment;
                        if (cmt != null)
                        {
                            var cur = SafeCommentText(cmt);
                            if (string.Equals(cur, text, StringComparison.Ordinal)) continue;

                            try { cmt.Text(text); }
                            catch
                            {
                                try { cmt.Delete(); } catch { }
                                try { cell.AddComment(text); } catch { /* ignore */ }
                            }
                        }
                        else
                        {
                            try { cell.AddComment(text); } catch { /* ignore */ }
                        }
                    }
                    finally { XqlCommon.ReleaseCom(cmt, cell); }
                }
            }
            finally
            {
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
                if (XqlSheet.TryGetHeaderMarker(ws, out marker))
                {
                    inter = ws.Application.Intersect(marker, target);
                    isHeaderEdit = inter != null;
                }

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

                _ = XqlCommon.OnExcelThreadAsync(() =>
                {
                    Excel.Worksheet? ws2 = null; Excel.Range? header2 = null;
                    try
                    {
                        var app2 = (Excel.Application)ExcelDnaUtil.Application;
                        ws2 = XqlSheet.FindWorksheet(app2, sheetName);
                        if (ws2 == null) return (object?)null;

                        if (!XqlSheet.TryGetHeaderMarker(ws2, out header2))
                            header2 = XqlSheet.GetHeaderRange(ws2);

                        var sm = sheetSvc.GetOrCreateSheet(sheetName);
                        ApplyHeaderUi(ws2, header2, sm, withValidation: true);
                        return (object?)null;
                    }
                    catch { return (object?)null; }
                    finally { XqlCommon.ReleaseCom(header2); XqlCommon.ReleaseCom(ws2); }
                });
            }
            finally
            {
                XqlCommon.ReleaseCom(inter, marker, lo);
            }
        }

        internal static void ApplyDataValidationForHeader(Excel.Worksheet ws, Excel.Range header, XqlSheet.Meta sm)
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

                        try { rng = lo.ListColumns[i]?.DataBodyRange; } catch { rng = null; }
                        if (rng == null) rng = ColBelowToEnd(ws, h);

                        // 🔒 KeyColumn(id) 잠금 + 입력 금지
                        if (!string.IsNullOrWhiteSpace(sm.KeyColumn) &&
                            string.Equals(name, sm.KeyColumn, StringComparison.OrdinalIgnoreCase))
                        {
                            LockIdColumn(ws, rng);
                            ApplyIdBlockedValidation(rng);
                        }
                        else
                        {
                            // 일반 컬럼은 타입별 DV
                            if (sm.Columns.TryGetValue(name!, out var ct))
                                ApplyValidationForKind(rng, ct.Kind);
                            else
                                try { rng.Validation.Delete(); } catch { /* clean only */ }
                            try { rng.Locked = false; } catch { }
                        }
                    }
                    finally { XqlCommon.ReleaseCom(h); XqlCommon.ReleaseCom(rng); }
                }
                // 시트 보호: UI에서만 잠금 적용(매크로는 허용)
                EnsureSheetProtectedUiOnly(ws);
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
                    if (!string.IsNullOrEmpty(sm.KeyColumn) &&
                        string.Equals(name, sm.KeyColumn, StringComparison.OrdinalIgnoreCase))
                    {
                        LockIdColumn(ws, col);
                        ApplyIdBlockedValidation(col);
                    }
                    else
                    {
                        if (!string.IsNullOrEmpty(name) && sm.Columns.TryGetValue(name!, out var ct))
                            ApplyValidationForKind(col, ct.Kind);
                        else
                            try { col.Validation.Delete(); } catch { /* clean only */ }
                        try { col.Locked = false; } catch { }
                    }
                }
                finally { XqlCommon.ReleaseCom(h, col); }
            }
            EnsureSheetProtectedUiOnly(ws);
        }

        // ─────────────────────────────────────────────────────────────────
        //  MarkTouchedCell / MarkInvalidCell ...
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
                Excel.Application app = (Excel.Application)ExcelDnaUtil.Application;
                Excel.Workbook? wb = null; Excel.Worksheet? ws = null;
                Excel.Range? r = null;
                try
                {
                    wb = app.ActiveWorkbook; if (wb == null) return (object?)null;
                    ws = FindOrCreateSheet(wb, "_XQL_Summary");

                    ws.Cells.ClearContents();
                    ws.Cells.ClearFormats();

                    int tables = _sumTables.Count;
                    double elapsedMs = TicksToMs(System.Diagnostics.Stopwatch.GetTimestamp() - _sumStartTicks);

                    Put(ws, 1, 1, title!, bold: true, size: 16);
                    Put(ws, 3, 1, "Tables"); Put(ws, 3, 2, tables.ToString());
                    Put(ws, 4, 1, "Batches"); Put(ws, 4, 2, _sumBatches.ToString());
                    Put(ws, 5, 1, "Affected Rows"); Put(ws, 5, 2, _sumAffected.ToString());
                    Put(ws, 6, 1, "Conflicts"); Put(ws, 6, 2, _sumConflicts.ToString());
                    Put(ws, 7, 1, "Errors"); Put(ws, 7, 2, _sumErrors.ToString());
                    Put(ws, 8, 1, "Elapsed (ms)"); Put(ws, 8, 2, elapsedMs.ToString("0"));

                    var box = ws.Range[ws.Cells[1, 1], ws.Cells[9, 3]];
                    try
                    {
                        var interior = box.Interior;
                        interior.Pattern = Excel.XlPattern.xlPatternSolid;
                        interior.Color = 0x00F0F0F0;
                        box.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    }
                    catch { }
                    finally { XqlCommon.ReleaseCom(box); }

#pragma warning disable CS8602
                    (ws.Columns["A:C"] as Excel.Range).AutoFit();
#pragma warning restore CS8602
                }
                catch { }
                finally { XqlCommon.ReleaseCom(r); XqlCommon.ReleaseCom(ws); XqlCommon.ReleaseCom(wb); }

                return (object?)null;

                static void Put(Excel.Worksheet w, int r0, int c0, string text, bool bold = false, int? size = null)
                {
                    var cell = (Excel.Range)w.Cells[r0, c0];
                    try
                    {
                        cell.Value2 = text;
                        if (bold) cell.Font.Bold = true;
                        if (size.HasValue) cell.Font.Size = size.Value;
                    }
                    finally { XqlCommon.ReleaseCom(cell); }
                }

                static double TicksToMs(long ticks)
                {
                    double freq = System.Diagnostics.Stopwatch.Frequency;
                    return ticks * 1000.0 / freq;
                }
            });
        }

        // ─────────────────────────────────────────────────────────────
        // Conflict 워크시트에 행 추가
        // ─────────────────────────────────────────────────────────────
        public static void AppendConflicts(IEnumerable<object>? conflicts)
        {
            if (conflicts == null) return;
            var items = conflicts.ToList();
            if (items.Count == 0) return;

            _ = XqlCommon.OnExcelThreadAsync(() =>
            {
                Excel.Application app = (Excel.Application)ExcelDnaUtil.Application;
                Excel.Workbook? wb = null; Excel.Worksheet? ws = null;
                Excel.Range? ur = null, row = null;
                try
                {
                    wb = app.ActiveWorkbook;
                    if (wb == null) return (object?)null;
                    ws = FindOrCreateSheet(wb, "_XQL_Conflicts");

                    ur = ws.UsedRange as Excel.Range;
                    bool needHeader = (ur?.Cells?.Count ?? 0) <= 1 || ((ws.Cells[1, 1] as Excel.Range)?.Value2 == null);
                    XqlCommon.ReleaseCom(ur); ur = null;
                    if (needHeader)
                    {
                        string[] headers = { "Timestamp", "Table", "RowKey", "Column", "Local", "Server", "Type", "Message", "Sheet", "Address" };
                        for (int i = 0; i < headers.Length; i++)
                            (ws.Cells[1, i + 1] as Excel.Range)!.Value2 = headers[i];
                        Excel.Range hdr = ws.Range[ws.Cells[1, 1], ws.Cells[1, headers.Length]];
                        try { ws.ListObjects.Add(Excel.XlListObjectSourceType.xlSrcRange, hdr, Type.Missing, Excel.XlYesNoGuess.xlYes); } catch { }
                        XqlCommon.ReleaseCom(hdr);
                    }

                    ur = ws.UsedRange as Excel.Range;
                    int last = (ur?.Row ?? 1) + ((ur?.Rows?.Count ?? 1) - 1);
                    XqlCommon.ReleaseCom(ur); ur = null;

                    foreach (var cf in items)
                    {
                        int next = Math.Max(2, last + 1);
                        row = ws.Range[ws.Cells[next, 1], ws.Cells[next, 10]];

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

#pragma warning disable CS8602
                        (row.Cells[1, 1] as Excel.Range).Value2 = ts;
                        (row.Cells[1, 2] as Excel.Range).Value2 = tbl;
                        (row.Cells[1, 3] as Excel.Range).Value2 = rk;
                        (row.Cells[1, 4] as Excel.Range).Value2 = col;
                        (row.Cells[1, 5] as Excel.Range).Value2 = loc;
                        (row.Cells[1, 6] as Excel.Range).Value2 = srv;
                        (row.Cells[1, 7] as Excel.Range).Value2 = typ;
                        (row.Cells[1, 8] as Excel.Range).Value2 = msg;
                        (row.Cells[1, 9] as Excel.Range).Value2 = sh;
                        (row.Cells[1, 10] as Excel.Range).Value2 = addr;
#pragma warning restore CS8602

                        try
                        {
                            var interior = row.Interior;
                            interior.Pattern = Excel.XlPattern.xlPatternSolid;
                            interior.Color = 0x00CCCCFF;
                        }
                        catch { }

                        if (!string.IsNullOrWhiteSpace(sh) && !string.IsNullOrWhiteSpace(addr))
                        {
                            try
                            {
                                string subAddr = $"'{sh.Replace("'", "''")}'!{addr}";
                                ws.Hyperlinks.Add(Anchor: row.Cells[1, 10], Address: "", SubAddress: subAddr, TextToDisplay: addr);
                            }
                            catch { }
                        }

                        last = next;
                        XqlCommon.ReleaseCom(row); row = null;
                    }
                }
                catch { }
                finally
                {
                    XqlCommon.ReleaseCom(row); XqlCommon.ReleaseCom(ur); XqlCommon.ReleaseCom(ws); XqlCommon.ReleaseCom(wb);
                }

                return (object?)null;

                static string Prop(object o, string name)
                    => Convert.ToString(PropObj(o, name), CultureInfo.InvariantCulture) ?? "";
                static object? PropObj(object o, string name)
                    => o.GetType().GetProperty(name, BindingFlags.Public | BindingFlags.Instance | BindingFlags.IgnoreCase)?.GetValue(o);
                static string ToStr(object? v)
                    => Convert.ToString(v, CultureInfo.InvariantCulture) ?? "";
            });
        }

        public readonly struct ExcelBatchScope : IDisposable
        {
            private readonly Excel.Application? _app;
            private readonly bool _oldEvents, _oldScreen, _oldAlerts;
            private readonly Excel.XlCalculation _oldCalc;

            public ExcelBatchScope(Excel.Application? app)
            {
                _app = app;
                if (app == null)
                {
                    _oldEvents = _oldScreen = _oldAlerts = false;
                    _oldCalc = Excel.XlCalculation.xlCalculationAutomatic;
                    return;
                }
                try
                {
                    _oldEvents = app.EnableEvents;
                    _oldScreen = app.ScreenUpdating;
                    _oldAlerts = app.DisplayAlerts;
                    _oldCalc = app.Calculation;

                    app.EnableEvents = false;
                    app.ScreenUpdating = false;
                    app.DisplayAlerts = false;
                    app.Calculation = Excel.XlCalculation.xlCalculationManual;
                }
                catch { /* ignore */ }
            }

            public void Dispose()
            {
                if (_app == null) return;
                try
                {
                    _app.Calculation = _oldCalc;
                    _app.DisplayAlerts = _oldAlerts;
                    _app.ScreenUpdating = _oldScreen;
                    _app.EnableEvents = _oldEvents;
                }
                catch { /* ignore */ }
            }
        }

        // ───────────────────────── 통합 엔트리: 헤더(plan) + 행 패치(patches) 한 번에 적용
        public static void ApplyPlanAndPatches(
            IReadOnlyDictionary<string, List<string>>? plan,
            IReadOnlyList<RowPatch>? patches)
        {
            if ((plan == null || plan.Count == 0) && (patches == null || patches.Count == 0))
                return;

            _ = XqlCommon.OnExcelThreadAsync(() =>
            {
                var app = (Excel.Application)ExcelDnaUtil.Application;
                if (app == null) return (object?)null;
                using var _ = new ExcelBatchScope(app);

                if (plan is { Count: > 0 })
                {
                    foreach (var (table, cols0) in plan)
                    {
                        if (string.IsNullOrWhiteSpace(table)) continue;
                        var cols = cols0?.Where(s => !string.IsNullOrWhiteSpace(s)).ToList() ?? new List<string>();
                        if (cols.Count == 0) continue;

                        EnsureHeaderForTable(app, table, cols);
                    }
                }

                if (patches is { Count: > 0 })
                {
                    InternalApplyCore(app, patches);
                    AppendFingerprintsForPatches(app, patches);
                }

                return (object?)null;
            });
        }

        // ───────────────────────── 패치 엔진
        private static void InternalApplyCore(Excel.Application app, IReadOnlyList<RowPatch> patches)
        {
            foreach (var grp in patches.GroupBy(p => p.Table, StringComparer.Ordinal))
            {
                Excel.Worksheet? ws = null; Excel.Range? header = null; Excel.ListObject? lo = null;
                try
                {
                    var smeta = default(XqlSheet.Meta);
                    ws = XqlSheet.FindWorksheetByTable(app, grp.Key, out smeta);
                    if (ws == null || smeta == null) continue;

                    lo = XqlSheet.FindListObjectByTable(ws, grp.Key);
                    header = lo?.HeaderRowRange ?? XqlSheet.GetHeaderRange(ws);
                    if (header == null) continue;

                    var headers = XqlSheet.ComputeHeaderNames(header);

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
                        XqlSheet.IsFallbackLetterHeader(header) ||
                        !headers.Any(h => serverCols.Contains(h));

                    if (needCreateHeader && serverCols.Count > 0)
                    {
                        var ordered = new List<string>(serverCols.Count);
                        if (serverCols.Contains(keyName)) ordered.Add(keyName);
                        ordered.AddRange(serverCols.Where(c => !string.Equals(c, keyName, StringComparison.Ordinal))
                                                   .OrderBy(c => c, StringComparer.Ordinal));

                        header = UpdateHeaderToColumns(ws, header, smeta, grp.Key, ordered);
                        headers = ordered;
                    }
                    if (headers.Count == 0) continue;

                    try { XqlAddIn.Sheet!.EnsureColumns(ws.Name, serverCols.ToArray()); } catch { }

                    int keyIdx1 = XqlSheet.FindKeyColumnIndex(headers, smeta.KeyColumn); // 1-based
                    int keyAbsCol = header.Column + keyIdx1 - 1;
                    int firstDataRow = header.Row + 1;

                    foreach (var patch in grp)
                    {
                        try
                        {
                            int? row = XqlSheet.FindRowByKey(ws, firstDataRow, keyAbsCol, patch.RowKey);
                            if (patch.Deleted)
                            {
                                if (row.HasValue) SafeDeleteRow(ws, row.Value);
                                continue;
                            }
                            if (!row.HasValue) row = AppendNewRow(ws, firstDataRow, lo);

                            ApplyCells(ws, row!.Value, header, headers, smeta, patch.Cells);
                        }
                        catch { /* per-row 안전 */ }
                    }
                }
                finally
                {
                    XqlCommon.ReleaseCom(lo, header, ws);
                }
            }
        }

        // ───────────────────────── 공용 헬퍼

        private static Excel.Range UpdateHeaderToColumns(
             Excel.Worksheet ws,
             Excel.Range oldHeader,
             XqlSheet.Meta smeta,
             string tableName,
             IList<string> columns)
        {
            // ✅ key(id) 1열 보장
            string keyName = string.IsNullOrWhiteSpace(smeta.KeyColumn) ? "id" : smeta.KeyColumn!;
            var cols = new List<string>(columns.Count + 1);
            if (!columns.Any(c => c.Equals(keyName, StringComparison.OrdinalIgnoreCase)))
                cols.Add(keyName);
            cols.AddRange(columns.Where(c => !c.Equals(keyName, StringComparison.OrdinalIgnoreCase)));

            // 새 헤더 영역 결정(1행, cols.Count 너비)
            var start = (Excel.Range)ws.Cells[oldHeader.Row, oldHeader.Column];
            var end = (Excel.Range)ws.Cells[oldHeader.Row, oldHeader.Column + cols.Count - 1];
            var newHeader = ws.Range[start, end];
            XqlCommon.ReleaseCom(start, end);

            // 값 채우기(배열 한 번에)
            var arr = new object[1, cols.Count];
            for (int i = 0; i < cols.Count; i++) arr[0, i] = cols[i] ?? "";
            newHeader.Value2 = arr;

            // 메타/마커/UI 동기화
            smeta.KeyColumn = keyName;
            XqlAddIn.Sheet!.EnsureColumns(ws.Name, cols);
            XqlSheet.SetHeaderMarker(ws, newHeader);
            ApplyHeaderUi(ws, newHeader, smeta, withValidation: true);
            InvalidateHeaderCache(ws.Name);
            RegisterTableSheet(tableName, ws.Name);

            return newHeader;
        }

        private static void EnsureHeaderForTable(Excel.Application app, string table, List<string> columns)
        {
            Excel.Worksheet? ws = null; Excel.Range? header = null;
            try
            {
                ws = XqlSheet.FindWorksheet(app, table) ?? app.ActiveSheet as Excel.Worksheet;
                if (ws == null)
                {
                    var sheets = app.Worksheets;
                    var last = (Excel.Worksheet)sheets[sheets.Count];
                    ws = (Excel.Worksheet)sheets.Add(After: last);
                    try { ws.Name = table; } catch { }
                    XqlCommon.ReleaseCom(last, sheets);
                }

                header = XqlSheet.GetHeaderRange(ws);
                var sm = XqlAddIn.Sheet!.GetOrCreateSheet(ws.Name);

                var curr = XqlSheet.ComputeHeaderNames(header);
                if (curr.Count != columns.Count || !curr.SequenceEqual(columns))
                {
                    header = UpdateHeaderToColumns(ws, header, sm, table, columns);
                }
                else
                {
                    XqlAddIn.Sheet!.EnsureColumns(ws.Name, columns);
                    XqlSheet.SetHeaderMarker(ws, header);
                    ApplyHeaderUi(ws, header, sm, withValidation: true);
                    RegisterTableSheet(table, ws.Name); // 🔧 FIX: sm.TableName → table
                }
            }
            finally { XqlCommon.ReleaseCom(header, ws); }
        }

        private static void AppendFingerprintsForPatches(Excel.Application app, IReadOnlyList<RowPatch> patches)
        {
            try
            {
                var wb = app.ActiveWorkbook;
                var items = new List<(string table, string rowKey, string colUid, string fp)>(Math.Max(64, patches.Count));

                foreach (var grp in patches.GroupBy(p => p.Table, StringComparer.Ordinal))
                {
                    Excel.Worksheet? ws = null; Excel.Range? header = null; Excel.ListObject? lo = null;
                    try
                    {
                        var smeta = default(XqlSheet.Meta);
                        ws = XqlSheet.FindWorksheetByTable(app, grp.Key, out smeta);
                        if (ws == null) continue;

                        lo = XqlSheet.FindListObjectByTable(ws, grp.Key);
                        header = lo?.HeaderRowRange ?? XqlSheet.GetHeaderRange(ws);
                        if (header == null) continue;

                        var uidMap = GetUidMapCached(ws, header);

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
                    finally { XqlCommon.ReleaseCom(lo, header, ws); }
                }

                if (items.Count > 0) XqlSheet.ShadowAppendFingerprints(wb, items);
            }
            catch { /* 무음 */ }
        }

        // XqlSheet에서 캐시를 활용할 수 있게 얇은 접근자 제공
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
                    var lr = lo.ListRows.Add(); XqlCommon.ReleaseCom(lr);
                    var body = lo.DataBodyRange;
                    if (body != null)
                    {
                        int row = body.Row + body.Rows.Count - 1;
                        XqlCommon.ReleaseCom(body);
                        return row;
                    }
                }
                catch { /* 폴백 */ }
            }
            int last = firstDataRow;
            try { var used = ws.UsedRange; last = used.Row + used.Rows.Count - 1; XqlCommon.ReleaseCom(used); }
            catch { }
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

                Excel.Range? rg = null;
                try
                {
                    rg = (Excel.Range)ws.Cells[row, header.Column + c];
                    if (val == null) { rg.Value2 = null; continue; }

                    switch (val)
                    {
                        case bool b: rg.Value2 = b; break;
                        case long l: rg.Value2 = (double)l; break;
                        case int i: rg.Value2 = (double)i; break;
                        case double d: rg.Value2 = d; break;
                        case float f: rg.Value2 = (double)f; break;
                        case decimal m: rg.Value2 = (double)m; break;
                        case DateTime dt: rg.Value2 = dt.ToOADate(); break;
                        default:
                            rg.Value2 = Convert.ToString(val, System.Globalization.CultureInfo.InvariantCulture);
                            break;
                    }

                    MarkTouchedCell(rg);
                }
                catch (Exception ex)
                {
                    XqlLog.Error($"패치 적용 실패: {ex.Message}", ws.Name, rg?.Address[false, false] ?? "");
                }
                finally { XqlCommon.ReleaseCom(rg); }
            }
        }

        private static void SafeDeleteRow(Excel.Worksheet ws, int row)
        {
            try { var rg = (Excel.Range)ws.Rows[row]; rg.Delete(); XqlCommon.ReleaseCom(rg); }
            catch { }
        }

        private static Excel.Worksheet FindOrCreateSheet(Excel.Workbook wb, string name)
        {
            foreach (Excel.Worksheet s in wb.Worksheets)
            {
                try { if (string.Equals(s.Name, name, StringComparison.Ordinal)) return s; }
                finally { XqlCommon.ReleaseCom(s); }
            }
            var created = (Excel.Worksheet)wb.Worksheets.Add();
            created.Name = name;
            created.Move(After: wb.Worksheets[wb.Worksheets.Count]);
            return created;
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

        private static void ClearHeaderUi(Excel.Worksheet ws, Excel.Range? header, bool removeMarker = false)
        {
            if (header == null) header = XqlSheet.GetHeaderRange(ws);

            foreach (Excel.Range cell in header.Cells)
            {
                try
                {
                    try { cell.ClearComments(); } catch { try { cell.Comment?.Delete(); } catch { } }
                }
                finally { XqlCommon.ReleaseCom(cell); }
            }

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

                        foreach (Excel.Range c in col.Cells)
                        {
                            try { XqlSheetView.TryClearInvalidMark(c); XqlSheetView.TryClearTouchedMark(c); }
                            finally { XqlCommon.ReleaseCom(c); }
                        }

                        XqlCommon.ReleaseCom(first, last);
                    }
                    finally { XqlCommon.ReleaseCom(col, h); }
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
            XqlCommon.ReleaseCom(first, last);
            return rng;
        }

        private static void ApplyValidationForKind(Excel.Range rng, XqlSheet.ColumnKind kind)
        {
            Excel.Validation? v = null;
            try
            {
                try
                {
                    if (rng == null) return;
                    if (rng.Areas != null && rng.Areas.Count > 1) return;
                    if ((long)rng.CountLarge == 0) return;
                }
                catch { /* ignore */ }

                try { rng.Validation.Delete(); } catch { }

                v = rng.Validation;

                bool added = false;

                string listSep = ",";
                try
                {
                    var app = (Excel.Application)rng.Application;
                    var sepObj = app.International[Excel.XlApplicationInternational.xlListSeparator];
                    if (sepObj is string s && !string.IsNullOrEmpty(s)) listSep = s;
                }
                catch { /* fallback , */ }

                switch (kind)
                {
                    case XqlSheet.ColumnKind.Int:
                        v.Add(
                            Excel.XlDVType.xlValidateWholeNumber,
                            Excel.XlDVAlertStyle.xlValidAlertStop,
                            Excel.XlFormatConditionOperator.xlBetween,
                            "-2147483648", "2147483647");
                        v.IgnoreBlank = true;
                        v.ErrorTitle = "정수만 허용";
                        v.ErrorMessage = "이 열은 정수만 입력할 수 있습니다.";
                        added = true;
                        break;

                    case XqlSheet.ColumnKind.Real:
                        v.Add(
                            Excel.XlDVType.xlValidateDecimal,
                            Excel.XlDVAlertStyle.xlValidAlertStop,
                            Excel.XlFormatConditionOperator.xlBetween,
                            "=-1E+307", "=1E+307");
                        v.IgnoreBlank = true;
                        v.ErrorTitle = "실수만 허용";
                        v.ErrorMessage = "이 열은 실수/숫자만 입력할 수 있습니다.";
                        added = true;
                        break;

                    case XqlSheet.ColumnKind.Bool:
                        v.Add(
                            Excel.XlDVType.xlValidateList,
                            Excel.XlDVAlertStyle.xlValidAlertStop,
                            Type.Missing, $"TRUE{listSep}FALSE", Type.Missing);
                        v.IgnoreBlank = true;
                        v.ErrorTitle = "BOOL만 허용";
                        v.ErrorMessage = "이 열은 TRUE 또는 FALSE만 입력할 수 있습니다.";
                        added = true;
                        break;

                    case XqlSheet.ColumnKind.Date:
                        var dmin = new DateTime(1900, 1, 1);
                        var dmax = new DateTime(9999, 12, 31);
                        v.Add(
                            Excel.XlDVType.xlValidateDate,
                            Excel.XlDVAlertStyle.xlValidAlertStop,
                            Excel.XlFormatConditionOperator.xlBetween,
                            dmin, dmax);
                        v.IgnoreBlank = true;
                        v.ErrorTitle = "날짜만 허용";
                        v.ErrorMessage = "이 열은 날짜 형식만 입력할 수 있습니다.";
                        added = true;
                        break;

                    default:
                        added = false;
                        break;
                }

                if (added)
                {
                    try { v.ShowError = true; } catch { }
                    try { v.ShowInput = false; } catch { }
                }
            }
            catch
            {
            }
            finally
            {
                XqlCommon.ReleaseCom(v);
            }
        }

        internal static Excel.Range? GetHeaderOrFallback(Excel.Worksheet ws)
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

        internal static void ApplyHeaderUi(Excel.Worksheet ws, Excel.Range header, XqlSheet.Meta sm, bool withValidation)
        {
            if (ws == null || header == null || sm == null) return;

            var tips = BuildHeaderTooltips(sm, header);
            SetHeaderTooltips(header, tips);
            ApplyHeaderOutlineBorder(header);

            if (withValidation)
                ApplyDataValidationForHeader(ws, header, sm);
        }

        /// <summary>
        /// 헤더/데이터를 실제로 재배치해서 'id'가 무조건 1열이 되도록 보장.
        /// 기존 id 값이 있다면 그대로 보존. 없으면 1열에 'id' 열을 새로 추가.
        /// </summary>
        private static Excel.Range EnsureIdIsFirst(Excel.Worksheet ws, Excel.Range header, XqlSheet.Meta sm, IList<string> names, bool reorderData)
        {
            string keyName = string.IsNullOrWhiteSpace(sm.KeyColumn) ? "id" : sm.KeyColumn!;
            int cols = header.Columns.Count;
            int idIdx = -1;
            for (int i = 0; i < names.Count; i++)
                if (string.Equals(names[i], keyName, StringComparison.OrdinalIgnoreCase)) { idIdx = i + 1; break; } // 1-based

            // 이미 1열이면 그대로
            if (idIdx == 1) return header;

            // 헤더에 없으면 1열에 새로 삽입
            if (idIdx < 0)
            {
                // 1) 헤더 왼쪽에 1열 삽입
                Excel.Range? firstCol = null;
                try
                {
                    firstCol = (Excel.Range)ws.Columns[header.Column]; // 헤더의 첫 Col
                    firstCol.Insert(Excel.XlInsertShiftDirection.xlShiftToRight);
                }
                catch { }
                finally { XqlCommon.ReleaseCom(firstCol); }

                // 2) 헤더 텍스트 'id'로 설정
                Excel.Range? idCell = null;
                try
                {
                    idCell = (Excel.Range)ws.Cells[header.Row, header.Column]; // 새 1열
                    idCell.Value2 = keyName;
                }
                finally { XqlCommon.ReleaseCom(idCell); }

                // 3) 헤더 Range 다시 계산(너비 +1)
                var start = (Excel.Range)ws.Cells[header.Row, header.Column];
                var end = (Excel.Range)ws.Cells[header.Row, header.Column + cols]; // +1
                var newHeader = ws.Range[start, end];
                XqlCommon.ReleaseCom(start, end);
                return newHeader;
            }

            // 헤더에 있으나 1열이 아니면: 열 전체를 앞으로 이동(값 보존)
            if (reorderData)
            {
                // idIdx(현재) → 1열로 이동
                Excel.Range? idWhole = null; Excel.Range? dest = null;
                try
                {
                    idWhole = ws.Range[
                    ws.Cells[header.Row, header.Column + (idIdx - 1)],
                    ws.Cells[ws.Rows.Count, header.Column + (idIdx - 1)]];

                    dest = (Excel.Range)ws.Cells[header.Row, header.Column];
                    idWhole.Cut(dest); // 앞으로 잘라붙이기
                }
                catch { /* 일부 환경에서 Cut이 막혀 있으면 포기(헤더명만 바꾸지 않음) */ }
                finally { XqlCommon.ReleaseCom(idWhole, dest); }
            }

            // 이동 후 헤더 범위 재계산
            int newCols = header.Columns.Count; // 동일
            var s = (Excel.Range)ws.Cells[header.Row, header.Column];
            var e = (Excel.Range)ws.Cells[header.Row, header.Column + newCols - 1];
            var hdr2 = ws.Range[s, e];
            XqlCommon.ReleaseCom(s, e);
            return hdr2;
        }

        /// <summary>id 컬럼을 잠그고( Locked=true ) 입력은 Custom Validation으로 차단</summary>
        private static void LockIdColumn(Excel.Worksheet ws, Excel.Range colData)
        {
            try
            {
                colData.Locked = true;
            }
            catch { }
        }

        /// <summary>Custom Validation으로 어떤 값도 허용하지 않음(=수정 불가). 빈 값은 그대로 둘 수 있게 하려면 필요시 수정.</summary>
        private static void ApplyIdBlockedValidation(Excel.Range rng)
        {
            Excel.Validation? v = null;
            try
            {
                try { rng.Validation.Delete(); } catch { }
                v = rng.Validation;
                // 항상 FALSE가 되도록: "=FALSE" → 사용자가 값을 입력하면 거부
                v.Add(Excel.XlDVType.xlValidateCustom, Excel.XlDVAlertStyle.xlValidAlertStop, Type.Missing, "=FALSE");
                v.ErrorTitle = "읽기 전용";
                v.ErrorMessage = "ID 열은 서버에서 관리됩니다.";
                v.ShowError = true;
                v.IgnoreBlank = true; // 비어있는 셀은 그대로 둘 수 있음
            }
            catch { }
            finally { XqlCommon.ReleaseCom(v); }
        }

        /// <summary>시트를 UI 한정으로 보호(UserInterfaceOnly=TRUE). 정렬/필터는 허용.</summary>
        private static void EnsureSheetProtectedUiOnly(Excel.Worksheet ws)
        {
            try
            {
                // 이미 보호 중이면 그대로 두되, 가능한 옵션만 보정
                bool protectedNow = false;
                try { protectedNow = ws.ProtectContents; } catch { }
                if (!protectedNow)
                {
                    ws.Protect(Password: Type.Missing, DrawingObjects: false, Contents: true, Scenarios: false,
                    UserInterfaceOnly: true, AllowFormattingCells: true, AllowFormattingColumns: true,
                    AllowFiltering: true, AllowSorting: true);
                }
            }
            catch { /* 보호 실패는 무시(회사 정책/공유통합문서 등) */ }
        }

        private static void EnqueueReapplyHeaderUi(string sheetName, bool withValidation)
        {
            string key = $"{sheetName}:{withValidation}";
            lock (_reapplyLock)
            {
                if (!_reapplyPending.Add(key)) return;
            }

            Task.Run(async () =>
            {
                await Task.Delay(150).ConfigureAwait(false);

                _ = XqlCommon.OnExcelThreadAsync(() =>
                {
                    Excel.Worksheet? ws2 = null; Excel.Range? h2 = null;
                    try
                    {
                        var app2 = (Excel.Application)ExcelDnaUtil.Application;
                        ws2 = XqlSheet.FindWorksheet(app2, sheetName);
                        if (ws2 == null) return (object?)null;

                        if (!XqlSheet.TryGetHeaderMarker(ws2, out h2))
                            h2 = XqlSheet.GetHeaderRange(ws2);
                        if (h2 == null) return (object?)null;

                        var sm = XqlAddIn.Sheet!.GetOrCreateSheet(sheetName);
                        ApplyHeaderUi(ws2, h2, sm, withValidation);
                        return (object?)null;
                    }
                    finally
                    {
                        XqlCommon.ReleaseCom(h2, ws2);
                        lock (_reapplyLock) { _reapplyPending.Remove(key); }
                    }
                });
            });
        }
    }
}
