// XqlSheetView.cs
using ExcelDna.Integration;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
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


        private static readonly object _sumLock = new();
        private static HashSet<string> _sumTables = new(StringComparer.Ordinal);
        private static int _sumAffected, _sumConflicts, _sumErrors, _sumBatches;
        private static long _sumStartTicks;


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

        // ─────────────────────────────────────────────────────────────────
        //  MarkTouchedCell: 서버 패치/중요 이벤트가 닿은 셀을 은은히 표시
        // ─────────────────────────────────────────────────────────────────
        public static void MarkTouchedCell(Excel.Range rg)
        {
            if (rg == null) return;
            try
            {
                var interior = rg.Interior;
                interior.Pattern = Excel.XlPattern.xlPatternSolid;
                // 연녹색 (0xCCFFCC) — 가독성 좋고 과하지 않음
                interior.Color = 0x00CCFFCC;
            }
            catch { /* ignore */ }
        }

        // 검증 실패 등 “주의” 셀 표시 (연한 붉은색)
        public static void MarkInvalidCell(Excel.Range rg)
        {
            if (rg == null) return;
            try
            {
                var interior = rg.Interior;
                interior.Pattern = Excel.XlPattern.xlPatternSolid;
                // 연분홍 (OLE BGR): 0xCCCCFF
                interior.Color = 0x00CCCCFF;
            }
            catch { /* ignore */ }
        }

        // === 새로 추가: 우리 마크만 조건부 해제 ===
        public static void TryClearInvalidMark(Excel.Range rg)
        {
            TryClearColor(rg, 0x00CCCCFF); // 연분홍
        }
        public static void TryClearTouchedMark(Excel.Range rg)
        {
            TryClearColor(rg, 0x00CCFFCC); // 연녹색
        }
        private static void TryClearColor(Excel.Range rg, int colorBgr)
        {
            if (rg == null) return;
            try
            {
                var it = rg.Interior;
                // Color는 Variant로 오므로 안전 변환
                int cur = Convert.ToInt32(it.Color);
                if (cur == colorBgr)
                    it.ColorIndex = Excel.XlColorIndex.xlColorIndexNone; // 사용자 색 보존
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
            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                Excel.Application app = (Excel.Application)ExcelDnaUtil.Application;
                Excel.Workbook? wb = null; Excel.Worksheet? ws = null;
                Excel.Range? r = null;
                try
                {
                    wb = app.ActiveWorkbook; if (wb == null) return;
                    ws = FindOrCreateSheet(wb, "_XQL_Summary");

                    // 시트 초기화(카드 영역만 깔끔하게)
                    ws.Cells.ClearContents();
                    ws.Cells.ClearFormats();

                    int tables = _sumTables.Count;
                    double elapsedMs = TicksToMs(System.Diagnostics.Stopwatch.GetTimestamp() - _sumStartTicks);

                    // 카드 렌더
                    Put(ws, 1, 1, title!, bold: true, size: 16);
                    Put(ws, 3, 1, "Tables");
                    Put(ws, 3, 2, tables.ToString());
                    Put(ws, 4, 1, "Batches");
                    Put(ws, 4, 2, _sumBatches.ToString());
                    Put(ws, 5, 1, "Affected Rows");
                    Put(ws, 5, 2, _sumAffected.ToString());
                    Put(ws, 6, 1, "Conflicts");
                    Put(ws, 6, 2, _sumConflicts.ToString());
                    Put(ws, 7, 1, "Errors");
                    Put(ws, 7, 2, _sumErrors.ToString());
                    Put(ws, 8, 1, "Elapsed (ms)");
                    Put(ws, 8, 2, elapsedMs.ToString("0"));

                    // 색상/강조
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

                    // 표준 컬럼 폭
#pragma warning disable CS8602 // null 가능 참조에 대한 역참조입니다.
                    (ws.Columns["A:C"] as Excel.Range).AutoFit();
#pragma warning restore CS8602 // null 가능 참조에 대한 역참조입니다.

                    // 내부 함수
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
                }
                catch { }
                finally { XqlCommon.ReleaseCom(r); XqlCommon.ReleaseCom(ws); XqlCommon.ReleaseCom(wb); }
            });

            static double TicksToMs(long ticks)
            {
                double freq = System.Diagnostics.Stopwatch.Frequency;
                return ticks * 1000.0 / freq;
            }
        }

        // ─────────────────────────────────────────────────────────────
        // Conflict 워크시트에 행 추가 (Conflicts shape가 달라도 Reflection로 안전파싱)
        // 컬럼: Timestamp | Table | RowKey | Column | Local | Server | Type | Message | Sheet | Address
        // ─────────────────────────────────────────────────────────────
        public static void AppendConflicts(IEnumerable<object>? conflicts)
        {
            if (conflicts == null) return;
            var items = conflicts.ToList();
            if (items.Count == 0) return;

            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                Excel.Application app = (Excel.Application)ExcelDnaUtil.Application;
                Excel.Workbook? wb = null; Excel.Worksheet? ws = null;
                Excel.Range? ur = null, row = null;
                try
                {
                    wb = app.ActiveWorkbook;
                    if (wb == null) return;
                    ws = FindOrCreateSheet(wb, "_XQL_Conflicts");

                    // 헤더 1회 보장
                    ur = ws.UsedRange as Excel.Range;
                    bool needHeader = (ur?.Cells?.Count ?? 0) <= 1 || ((ws.Cells[1, 1] as Excel.Range)?.Value2 == null);
                    XqlCommon.ReleaseCom(ur); ur = null;
                    if (needHeader)
                    {
                        string[] headers = { "Timestamp", "Table", "RowKey", "Column", "Local", "Server", "Type", "Message", "Sheet", "Address" };
                        for (int i = 0; i < headers.Length; i++)
                            (ws.Cells[1, i + 1] as Excel.Range)!.Value2 = headers[i];
                        // 간단 오토필터
                        Excel.Range hdr = ws.Range[ws.Cells[1, 1], ws.Cells[1, headers.Length]];
                        try { ws.ListObjects.Add(Excel.XlListObjectSourceType.xlSrcRange, hdr, Type.Missing, Excel.XlYesNoGuess.xlYes); } catch { }
                        XqlCommon.ReleaseCom(hdr);
                    }

                    // 현재 마지막 행
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

                        // 값 채우기
#pragma warning disable CS8602 // null 가능 참조에 대한 역참조입니다.
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
#pragma warning restore CS8602 // null 가능 참조에 대한 역참조입니다.

                        // 약한 색 (주의 = 연분홍)
                        try
                        {
                            var interior = row.Interior;
                            interior.Pattern = Excel.XlPattern.xlPatternSolid;
                            interior.Color = 0x00CCCCFF;
                        }
                        catch { }

                        // 대상 셀 하이퍼링크 (가능할 때)
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
            });

            // —— 로컬 헬퍼
            static string Prop(object o, string name)
                => Convert.ToString(PropObj(o, name), CultureInfo.InvariantCulture) ?? "";
            static object? PropObj(object o, string name)
                => o.GetType().GetProperty(name, BindingFlags.Public | BindingFlags.Instance | BindingFlags.IgnoreCase)?.GetValue(o);
            static string ToStr(object? v)
                => Convert.ToString(v, CultureInfo.InvariantCulture) ?? "";
        }


        // Private

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

                        // 우리 마크(연녹색/연분홍)만 조건부로 제거
                        foreach (Excel.Range c in col.Cells)
                        {
                            try { XqlSheetView.TryClearInvalidMark(c); XqlSheetView.TryClearTouchedMark(c); }
                            finally { XqlCommon.ReleaseCom(c); }
                        }

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
                // 빈/다중 영역 스킵 (DV 예외 방지)
                try
                {
                    if (rng == null) return;
                    if ((long)rng.CountLarge == 0) return;
                    if (rng.Areas != null && rng.Areas.Count > 1) return;
                }
                catch { /* ignore */ }

                // 기존 규칙 제거(잔존 규칙 때문에 Add 실패 방지)
                try { rng.Validation.Delete(); } catch { }

                v = rng.Validation;

                bool added = false;

                // 지역설정: 리스트 구분자(, 또는 ;)
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
                    case ColumnKind.Int:
                        // 안전한 32bit 정수 범위(문자열로 전달해도 OK)
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

                    case ColumnKind.Real:
                        // Excel이 싫어하는 ±1.79e308 대신 ±1e307로 클램프 (안전)
                        // 문자열로 전달(=수식) 대신 상수로도 되지만 로캘 영향 줄이려 문자열 사용
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

                    case ColumnKind.Bool:
                        // TRUE/FALSE 목록 — 로캘별 리스트 구분자 사용
                        v.Add(
                            Excel.XlDVType.xlValidateList,
                            Excel.XlDVAlertStyle.xlValidAlertStop,
                            Type.Missing, $"TRUE{listSep}FALSE", Type.Missing);
                        v.IgnoreBlank = true;
                        v.ErrorTitle = "BOOL만 허용";
                        v.ErrorMessage = "이 열은 TRUE 또는 FALSE만 입력할 수 있습니다.";
                        added = true;
                        break;

                    case ColumnKind.Date:
                        // 지역화된 함수명 문제 회피: DateTime 값을 직접 전달
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

                    // TEXT/JSON/ANY 등은 DV 미적용 (서버/런타임 검증으로 처리)
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
                // 병합/시트 보호/특수 범위 등으로 실패할 수 있음 — 조용히 무시
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
