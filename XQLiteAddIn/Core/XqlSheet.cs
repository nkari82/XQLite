// XqlSheet.cs
using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using static XQLite.AddIn.XqlCommon;
using Excel = Microsoft.Office.Interop.Excel;

namespace XQLite.AddIn
{
    /// <summary>
    /// 시트 범용 유틸(찾기/상대키/헤더 Fallback 등) + "메타 레지스트리" 기능을 흡수한 인스턴스 서비스.
    /// UI 로직은 XqlSheetView가 담당한다.
    /// </summary>
    internal sealed class XqlSheet
    {
        // ───────────────────────── Models
        internal sealed class Meta
        {
            public string TableName { get; set; } = "";
            public string KeyColumn { get; set; } = "id";
            public Dictionary<string, ColumnType> Columns { get; } = new(StringComparer.Ordinal);
            public void SetColumn(string name, ColumnType t) => Columns[name] = t;
        }

        internal enum ColumnKind { Text, Int, Real, Bool, Date, Json, Any }

        internal sealed class ColumnType
        {
            public ColumnKind Kind { get; set; } = ColumnKind.Text;
            public bool Nullable { get; set; } = true;

            public string ToTooltip()
            {
                var k = Kind switch
                {
                    ColumnKind.Int => "INT",
                    ColumnKind.Real => "REAL",
                    ColumnKind.Bool => "BOOL",
                    ColumnKind.Date => "DATE",
                    ColumnKind.Json => "JSON",
                    _ => "TEXT"
                };
                var nul = Nullable ? "NULL OK" : "NOT NULL";
                return $"{k} • {nul}";
            }
        }

        internal enum ErrCode
        {
            None = 0,
            E_TYPE_MISMATCH,
            E_RANGE,
            E_CHECK_FAIL,
            E_JSON_PARSE,
            E_NULL_NOT_ALLOWED,
            E_UNSUPPORTED,
        }

        internal readonly struct ValidationResult
        {
            public bool IsOk { get; }
            public ErrCode Code { get; }
            public string Message { get; }

            private ValidationResult(bool ok, ErrCode code, string msg)
            {
                IsOk = ok; Code = code; Message = msg;
            }

            public static ValidationResult Ok() => new(true, ErrCode.None, "");
            public static ValidationResult Fail(ErrCode code, string msg) => new(false, code, msg);
        }

        private readonly Dictionary<string, Meta> _sheets = new(StringComparer.Ordinal);
        private const string StateSheetName = ".xql_state";   // VeryHidden
        private const string ShadowSheetName = ".xql_shadow"; // VeryHidden

        // ───────────────────────── Meta registry
        public bool TryGetSheet(string sheetName, out Meta sm) => _sheets.TryGetValue(sheetName, out sm!);

        public Meta GetOrCreateSheet(string sheetName, string defaultKeyColumn = "id")
        {
            if (!_sheets.TryGetValue(sheetName, out var sm))
            {
                sm = new Meta { TableName = sheetName, KeyColumn = defaultKeyColumn };
                _sheets[sheetName] = sm;
            }
            return sm;
        }

        /// <summary>컬럼 목록을 받아 없으면 기본 타입으로 생성.</summary>
        public IReadOnlyList<string> EnsureColumns(string sheetName, IEnumerable<string> columnNames,
            ColumnKind defaultKind = ColumnKind.Text, bool defaultNullable = true)
        {
            var sm = GetOrCreateSheet(sheetName);
            var added = new List<string>();
            foreach (var raw in columnNames)
            {
                var name = (raw ?? "").Trim();
                if (string.IsNullOrEmpty(name)) continue;

                if (!sm.Columns.TryGetValue(name, out _))
                {
                    var ct = new ColumnType { Kind = defaultKind, Nullable = defaultNullable };
                    sm.SetColumn(name, ct);
                    added.Add(name);
                }
            }
            return added;
        }

        // ───────────────────────── UsedRange 안전 접근 유틸
        /// <summary>
        /// UsedRange의 경계(첫행/첫열/끝행/끝열)를 안전하게 얻는다.
        /// 실패 시 최소값(1,1,1,1)로 채우며 false 반환. 성공 시 true.
        /// </summary>
        internal static bool TryGetUsedBounds(Excel.Worksheet ws,
            out int firstRow, out int firstCol, out int lastRow, out int lastCol)
        {
            firstRow = 1; firstCol = 1; lastRow = 1; lastCol = 1;
            if (ws == null) return false;

            try
            {
                // 빠른 유효성 검사: ws.Name 접근
                try { _ = ws.Name; }
                catch (InvalidComObjectException) { return false; }
                catch (COMException) { return false; }

                // 1) UsedRange
                using (var used = SmartCom<Range>.Wrap(ws.UsedRange))
                {
                    if (used.Value != null)
                    {
                        try { var _ = used.Value.Address[true, true, Excel.XlReferenceStyle.xlA1, false]; } catch { }
                        firstRow = used.Value.Row;
                        firstCol = used.Value.Column;
                        lastRow = firstRow + used.Value.Rows.Count - 1;
                        lastCol = firstCol + used.Value.Columns.Count - 1;
                        if (lastRow < firstRow) lastRow = firstRow;
                        if (lastCol < firstCol) lastCol = firstCol;
                        return true;
                    }
                }

                // 2) SpecialCells(xlCellTypeLastCell)
                using (var lastCell = SmartCom<Range>.Wrap(ws.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell) as Excel.Range))
                {
                    if (lastCell.Value != null)
                    {
                        lastRow = lastCell.Value.Row;
                        lastCol = lastCell.Value.Column;
                        firstRow = 1; firstCol = 1;
                        if (lastRow < firstRow) lastRow = firstRow;
                        if (lastCol < firstCol) lastCol = firstCol;
                        return true;
                    }
                }

                // 3) Find("*", xlPrevious)
                using (var after = SmartCom<Range>.Wrap(ws.Cells[1, 1] as Excel.Range))
                using (var finder = SmartCom<Range>.Wrap(ws.Cells.Find(What: "*",
                        After: after.Value,
                        LookIn: Excel.XlFindLookIn.xlValues,
                        LookAt: Excel.XlLookAt.xlPart,
                        SearchOrder: Excel.XlSearchOrder.xlByRows,
                        SearchDirection: Excel.XlSearchDirection.xlPrevious,
                        MatchCase: false)))
                {
                    if (finder.Value != null)
                    {
                        lastRow = finder.Value.Row;
                        lastCol = finder.Value.Column;
                        firstRow = 1; firstCol = 1;
                        return true;
                    }
                }

                firstRow = firstCol = lastRow = lastCol = 1;
                return false;
            }
            catch { return false; }
        }

        // ───────────────────────── Header (1행)
        /// <summary>
        /// 1행 헤더 Range와 트림된 헤더 텍스트 리스트를 반환.
        /// 반환된 header는 호출자가 해제 책임(Detach)로 가져감.
        /// </summary>
        internal static (Excel.Range header, List<string> names) GetHeaderAndNames(Excel.Worksheet ws)
        {
            TryGetUsedBounds(ws, out var fr, out var fc, out var lr, out var lc);
            int firstCol = Math.Max(1, fc);
            int lastCol = Math.Max(firstCol, lc);

            using var s = SmartCom<Range>.Wrap(ws.Cells[1, firstCol] as Excel.Range);
            using var e = SmartCom<Range>.Wrap(ws.Cells[1, lastCol] as Excel.Range);
            using var header = SmartCom<Range>.Wrap(ws.Range[s.Value, e.Value]);
            var detached = header.Detach()!;
            var names = ComputeHeaderNames(detached);
            return (detached, names);
        }

        /// <summary>
        /// 헤더 Fallback: 1행 전체 중에서 마지막 실제 텍스트가 있는 열까지를 헤더로 간주.
        /// </summary>
        internal static Excel.Range GetHeaderRange(Excel.Worksheet sh)
        {
            TryGetUsedBounds(sh, out _, out _, out _, out var lc);
            int hr = 1;
            int last = 1;
            int lastCol = Math.Max(1, lc);

            for (int c = lastCol; c >= 1; --c)
            {
                using var cell = SmartCom<Range>.Wrap(sh.Cells[hr, c] as Excel.Range);
                string txt = (Convert.ToString(cell.Value?.Value2) ?? "").Trim();
                if (!string.IsNullOrEmpty(txt)) { last = c; break; }
            }

            using var s = SmartCom<Range>.Wrap(sh.Cells[hr, 1] as Excel.Range);
            using var e = SmartCom<Range>.Wrap(sh.Cells[hr, last] as Excel.Range);
            return SmartCom<Range>.Wrap(sh.Range[s.Value, e.Value]).Detach()!;
        }

        /// <summary>헤더 Range에서 트림된 컬럼명 배열을 뽑아낸다. 빈 값은 A/B/C…로 대체.</summary>
        internal static List<string> ComputeHeaderNames(Excel.Range header)
        {
            var names = new List<string>(Math.Max(1, header?.Columns.Count ?? 0));
            if (header == null) return names;

            var v = header.Value2 as object[,];
            if (v != null)
            {
                int cols = header.Columns.Count;
                for (int c = 1; c <= cols; c++)
                {
                    var raw = (Convert.ToString(v[1, c]) ?? string.Empty).Trim();
                    if (string.IsNullOrEmpty(raw))
                    {
                        using var hc = SmartCom<Range>.Wrap(header.Cells[1, c] as Excel.Range);
                        raw = XqlCommon.ColumnIndexToLetter((hc.Value as Range)!.Column);
                    }
                    names.Add(raw);
                }
                return names;
            }

            for (int i = 1; i <= header.Columns.Count; i++)
            {
                using var hc = SmartCom<Range>.Wrap(header.Cells[1, i] as Excel.Range);
                var raw = ((hc.Value as Range)!.Value2 as string)?.Trim();
                names.Add(string.IsNullOrEmpty(raw) ? XqlCommon.ColumnIndexToLetter(hc.Value.Column) : raw!);
            }
            return names;
        }

        /// <summary>헤더가 A/B/C… 기본 글자들로만 구성되어 있는지 판별.</summary>
        internal static bool IsFallbackLetterHeader(Excel.Range header)
        {
            if (header == null || header.Columns.Count == 0) return true;
            for (int i = 1; i <= header.Columns.Count; i++)
            {
                using var hc = SmartCom<Range>.Wrap(header.Cells[1, i] as Excel.Range);
                var name = ((hc.Value as Range)!.Value2 as string)?.Trim() ?? "";
                var expect = XqlCommon.ColumnIndexToLetter(header.Column + i - 1);
                if (!string.Equals(name, expect, StringComparison.Ordinal))
                    return false;
            }
            return true;
        }

        /// <summary>
        /// 테이블명으로 워크시트를 찾는다. (캐시 사용하지 않고 매번 통합문서 스캔)
        /// </summary>
        internal static Excel.Worksheet? FindWorksheetByTable(Excel.Application app, string table, out XqlSheet.Meta? smeta)
        {
            smeta = null;
            if (app == null) return null;

            // 1) 먼저 ActiveWorkbook에서만 시도(가장 안전)
            try
            {
                using var wbW = SmartCom<Excel.Workbook>.Wrap(app.ActiveWorkbook);
                if (wbW?.Value != null)
                {
                    var hit = FindInWorkbook(wbW.Value, table, out smeta);
                    if (hit != null) return hit;
                }
            }
            catch { /* ignore */ }

            // 2) 열려있는 모든 Workbook을 안전하게 순회
            try
            {
                using var wbsW = SmartCom<Excel.Workbooks>.Wrap(app.Workbooks);
                int wbc = wbsW?.Value?.Count ?? 0;
                for (int wi = 1; wi <= wbc; wi++)
                {
                    using var wb = SmartCom<Excel.Workbook>.Acquire(() => wbsW!.Value![wi]);
                    if (wb?.Value == null) continue;

                    var hit = FindInWorkbook(wb.Value, table, out smeta);
                    if (hit != null) return hit;
                }
            }
            catch { /* ignore */ }

            return null;
        }

        private static Excel.Worksheet? FindInWorkbook(Excel.Workbook wb, string table, out XqlSheet.Meta? smeta)
        {
            smeta = null;
            try
            {
                using var sheetsW = SmartCom<Excel.Sheets>.Wrap(wb.Worksheets);
                int count = sheetsW?.Value?.Count ?? 0;
                for (int i = 1; i <= count; i++)
                {
                    Excel.Worksheet? raw = null;
                    try { raw = sheetsW!.Value![i] as Excel.Worksheet; }
                    catch { continue; }

                    using var w = SmartCom<Excel.Worksheet>.Wrap(raw);
                    if (w?.Value == null) continue;

                    try
                    {
                        // 메타 → 테이블 이름 매핑 우선
                        if (XqlAddIn.Sheet!.TryGetSheet(w.Value.Name, out var m))
                        {
                            var t = string.IsNullOrWhiteSpace(m.TableName) ? w.Value.Name : m.TableName!;
                            if (string.Equals(t, table, StringComparison.OrdinalIgnoreCase))
                            {
                                var ret = w.Value; w.Detach(); smeta = m; return ret;
                            }
                        }
                        // 메타가 없으면 시트명 직접 매칭
                        else if (string.Equals(w.Value.Name, table, StringComparison.OrdinalIgnoreCase))
                        {
                            smeta = XqlAddIn.Sheet!.GetOrCreateSheet(w.Value.Name);
                            var ret = w.Value; w.Detach(); return ret;
                        }
                    }
                    catch { /* continue */ }
                }
            }
            catch { /* ignore */ }

            return null;
        }

        /// <summary>
        /// 이 시트가 부트스트랩(헤더/데이터 초기화)이 필요한 상태인지 판단.
        /// 마커 없음 OR 사실상 비어있음 OR 1행이 A/B/C…만 있을 때 true.
        /// </summary>
        internal static bool NeedsBootstrap(Excel.Worksheet ws)
        {
            try
            {
                if (TryGetHeaderMarker(ws, out var hdr)) { using var _ = SmartCom<Range>.Wrap(hdr); return false; }

                if (!TryGetUsedBounds(ws, out var fr, out var fc, out var lr, out var lc)) return false;
                if (fr == lr && fc == lc) return true; // 사실상 비어있음(1셀)

                using var header = SmartCom<Range>.Wrap(GetHeaderRange(ws));
                return IsFallbackLetterHeader((header.Value as Range)!);
            }
            catch { return false; }
        }

        // ───────────────────────── Finders (name/sheet/listobject)
        internal static Excel.Worksheet? FindWorksheet(Excel.Application app, string sheetName)
        {
            if (app is null || string.IsNullOrWhiteSpace(sheetName)) return null;

            Excel.Worksheet? match = null;
            foreach (Excel.Worksheet? ws in app.Worksheets)
            {
                if (ws is null) continue;
                if (string.Equals(ws.Name, sheetName, StringComparison.Ordinal))
                {
                    match = ws;
                    break;
                }
            }
            return match;
        }

        internal static Excel.ListObject? FindListObject(Excel.Worksheet ws, string listObjectName)
        {
            if (ws is null || string.IsNullOrWhiteSpace(listObjectName)) return null;

            Excel.ListObject? match = null;
            foreach (Excel.ListObject? lo in ws.ListObjects)
            {
                if (lo is null) continue;
                if (string.Equals(lo.Name, listObjectName, StringComparison.Ordinal))
                {
                    match = lo;
                    break;
                }
            }
            return match;
        }

        internal static Excel.ListObject? FindListObjectContaining(Excel.Worksheet ws, Excel.Range? rng)
        {
            if (ws is null || rng is null) return null;

            foreach (Excel.ListObject lo in ws.ListObjects)
            {
                using var header = SmartCom<Range>.Wrap(lo.HeaderRowRange);
                using var body = SmartCom<Range>.Wrap(lo.DataBodyRange);

                bool hit = false;
                if (header.Value != null)
                {
                    using var inter1 = SmartCom<Range>.Wrap(XqlCommon.IntersectSafe(ws, rng, header.Value));
                    hit |= inter1.Value != null;
                }
                if (!hit && body.Value != null)
                {
                    using var inter2 = SmartCom<Range>.Wrap(XqlCommon.IntersectSafe(ws, rng, body.Value));
                    hit |= inter2.Value != null;
                }
                if (hit) return lo;
            }
            return null;
        }

        internal static Excel.ListObject? FindListObjectByTable(Excel.Worksheet ws, string tableNameOrQualified)
        {
            if (ws is null || string.IsNullOrWhiteSpace(tableNameOrQualified)) return null;

            string? sheetHint = null;
            string loName = tableNameOrQualified;
            var dot = tableNameOrQualified.IndexOf('.');
            if (dot >= 0)
            {
                sheetHint = tableNameOrQualified.Substring(0, dot);
                loName = tableNameOrQualified.Substring(dot + 1);
            }

            var found = FindListObject(ws, loName) ?? FindListObjectByHeaderCaption(ws, loName);
            if (found != null) return found;

            if (!string.IsNullOrEmpty(sheetHint) && ws.Application is Excel.Application app)
            {
                var ws2 = FindWorksheet(app, sheetHint!);
                if (ws2 != null)
                {
                    var f2 = FindListObject(ws2, loName) ?? FindListObjectByHeaderCaption(ws2, loName);
                    if (f2 != null) return f2;
                }
            }

            foreach (Excel.ListObject? lo in ws.ListObjects)
            {
                if (lo is null) continue;
                var qualified = $"{ws.Name}.{lo.Name}";
                if (string.Equals(qualified, tableNameOrQualified, StringComparison.Ordinal))
                {
                    return lo;
                }
            }

            var app2 = ws.Application as Excel.Application;
            if (app2 != null)
            {
                foreach (Excel.Worksheet w in app2.Worksheets)
                {
                    foreach (Excel.ListObject lo2 in w.ListObjects)
                    {
                        if (string.Equals(lo2.Name, loName, StringComparison.Ordinal) ||
                            string.Equals($"{w.Name}.{lo2.Name}", tableNameOrQualified, StringComparison.Ordinal))
                            return lo2;
                    }
                }
            }
            return null;
        }

        private static Excel.ListObject? FindListObjectByHeaderCaption(Excel.Worksheet ws, string caption)
        {
            foreach (Excel.ListObject? lo in ws.ListObjects)
            {
                if (lo is null) continue;
                using var header = SmartCom<Range>.Wrap(lo.HeaderRowRange);
                if (header.Value == null) continue;

                var v = header.Value.Value2 as object[,];
                if (v == null) continue;

                string first = Convert.ToString(v[1, 1]) ?? string.Empty;
                if (string.Equals(first, caption, StringComparison.Ordinal))
                {
                    return lo;
                }

                int cols = header.Value.Columns.Count;
                var joined = string.Empty;
                for (int i = 1; i <= cols; i++)
                    joined += (i == 1 ? "" : "|") + (Convert.ToString(v[1, i]) ?? string.Empty);
                if (string.Equals(joined, caption, StringComparison.Ordinal))
                {
                    return lo;
                }
            }
            return null;
        }

        internal static int FindKeyColumnIndex(List<string> headers, string keyName)
        {
            if (!string.IsNullOrWhiteSpace(keyName))
            {
                var idx = headers.FindIndex(h => string.Equals(h, keyName, StringComparison.Ordinal));
                if (idx >= 0) return idx + 1;
            }
            var id = headers.FindIndex(h => string.Equals(h, "id", StringComparison.OrdinalIgnoreCase));
            if (id >= 0) return id + 1;
            var key = headers.FindIndex(h => string.Equals(h, "key", StringComparison.OrdinalIgnoreCase));
            if (key >= 0) return key + 1;
            return 1;
        }

        // ───────────────────────── Relative Keys
        internal static string ColumnKey(string sheet, string table, int hRow, int hCol, int colOffset, string? hdrName = null)
            => hdrName is { Length: > 0 }
               ? $"col:{sheet}:{table}:H{hRow}C{hCol}:dx={colOffset}:hdr={Escape(hdrName)}"
               : $"col:{sheet}:{table}:H{hRow}C{hCol}:dx={colOffset}";

        internal static string CellKey(string sheet, string table, int hRow, int hCol, int rowOffset, int colOffset)
            => $"cell:{sheet}:{table}:H{hRow}C{hCol}:dr={rowOffset}:dc={colOffset}";

        internal readonly record struct RelDesc(string Kind, string Sheet, string Table, int AnchorRow, int AnchorCol, int RowOffset, int ColOffset, string? HeaderNameHint);

        internal static bool TryParse(string key, out RelDesc d)
        {
            d = default;
            try
            {
                var parts = key.Split(':');
                if (parts.Length < 4) return false;
                var kind = parts[0];
                var sheet = parts[1];
                var table = parts[2];

                int hRow = 0, hCol = 0, dr = 0, dc = 0;
                string? hdr = null;

                foreach (var seg in parts.Skip(3))
                {
                    if (seg.StartsWith("H") && seg.Contains('C'))
                    {
                        var hc = seg.Substring(1).Split('C');
                        hRow = int.Parse(hc[0]); hCol = int.Parse(hc[1]);
                    }
                    else if (seg.StartsWith("dx=")) dc = ParseIntOrDefault(seg, 3);
                    else if (seg.StartsWith("dr=")) dr = ParseIntOrDefault(seg, 3);
                    else if (seg.StartsWith("dc=")) dc = ParseIntOrDefault(seg, 3);
                    else if (seg.StartsWith("hdr=")) hdr = Unescape(seg.Substring(4));
                }

                d = new RelDesc(kind, sheet, table, hRow, hCol, dr, dc, hdr);
                return true;
            }
            catch { return false; }
        }

        /// <summary>
        /// 리소스 키(col/cell)를 Range로 해석. ListObject가 없으면 헤더 마커 → 폴백 헤더로 앵커를 찾는다.
        /// </summary>
        internal static bool TryResolve(
            Excel.Application app,
            RelDesc d,
            out Excel.Range? target,
            out int? targetHeaderCol,
            out Excel.ListObject? lo)
        {
            target = null; targetHeaderCol = null; lo = null;

            try
            {
                using var _ws = SmartCom<Excel.Worksheet>.Wrap(FindWorksheet(app, d.Sheet));
                if (_ws.Value == null) return false;

                // 1) ListObject 있으면 우선 사용
                using var _lo = SmartCom<Excel.ListObject>.Wrap(
                    FindListObject(_ws.Value, d.Table) ?? FindListObjectByTable(_ws.Value, d.Table));

                int anchorRow, anchorCol;

                if (_lo.Value != null && _lo.Value.HeaderRowRange != null)
                {
                    using var _hdr = SmartCom<Excel.Range>.Wrap(_lo.Value.HeaderRowRange);
                    anchorRow = _hdr.Value!.Row;
                    anchorCol = _hdr.Value!.Column;
                    lo = _lo.Detach();
                }
                else
                {
                    // 2) 헤더 마커 → 3) 폴백 헤더
                    Excel.Range? hdr = null;
                    if (TryGetHeaderMarker(_ws.Value, out var m)) hdr = m;
                    hdr ??= GetHeaderRange(_ws.Value);

                    using var _hdr = SmartCom<Excel.Range>.Wrap(hdr);
                    if (_hdr.Value == null) return false;

                    anchorRow = _hdr.Value.Row;
                    anchorCol = _hdr.Value.Column;
                }

                if (string.Equals(d.Kind, "col", StringComparison.OrdinalIgnoreCase))
                {
                    int col = anchorCol + d.ColOffset;
                    targetHeaderCol = col;

                    using var _cell = SmartCom<Excel.Range>.Acquire(() =>
                        (Excel.Range)_ws.Value!.Cells[anchorRow, col]);

                    target = _cell.Detach();
                    return target != null;
                }
                else if (string.Equals(d.Kind, "cell", StringComparison.OrdinalIgnoreCase))
                {
                    int row = (anchorRow + 1) + d.RowOffset;   // 데이터 첫 행 = header+1
                    int col = anchorCol + d.ColOffset;

                    using var _cell = SmartCom<Excel.Range>.Acquire(() =>
                        (Excel.Range)_ws.Value!.Cells[row, col]);

                    target = _cell.Detach();
                    return target != null;
                }

                return false;
            }
            catch
            {
                return false;
            }
        }

        // ───────────────────────── 기타 유틸
        public static int? FindRowByKey(Excel.Worksheet ws, int firstDataRow, int keyAbsCol, object rowKey)
        {
            try
            {
                using var c1 = SmartCom<Range>.Wrap(ws.Cells[firstDataRow, keyAbsCol]);
                using var c2 = SmartCom<Range>.Wrap(ws.Cells[ws.Rows.Count, keyAbsCol]);
                using var col = SmartCom<Range>.Wrap(ws.Range[c1.Value, c2.Value]);

                using var hit = SmartCom<Range>.Wrap((col.Value as Range)!.Find(
                    What: rowKey,
                    LookIn: Excel.XlFindLookIn.xlValues,
                    LookAt: Excel.XlLookAt.xlWhole,
                    SearchOrder: Excel.XlSearchOrder.xlByRows,
                    SearchDirection: Excel.XlSearchDirection.xlNext,
                    MatchCase: false));

                return hit.Value?.Row;
            }
            catch { return null; }
        }

        /// <summary>
        /// 셀 값 검증(형/제약). 존재하지 않는 시트/컬럼이면 통과(보수적으로 허용).
        /// </summary>
        internal ValidationResult ValidateCell(string sheet, string col, object? value)
        {
            if (!_sheets.TryGetValue(sheet, out var sm))
                return ValidationResult.Ok();

            if (!sm.Columns.TryGetValue(col, out var ct))
                return ValidationResult.Ok();

            // NotNull
            if (!ct.Nullable && XqlCommon.IsNullish(value))
                return ValidationResult.Fail(ErrCode.E_NULL_NOT_ALLOWED, "Null/empty is not allowed.");

            // 타입별 검증
            switch (ct.Kind)
            {
                case ColumnKind.Int:
                    {
                        if (XqlCommon.IsNullish(value))
                            return ValidationResult.Ok();
                        if (!XqlCommon.TryToInt64(value!, out _))
                            return ValidationResult.Fail(ErrCode.E_TYPE_MISMATCH, "Expect INT.");
                        break;
                    }
                case ColumnKind.Real:
                    {
                        if (XqlCommon.IsNullish(value)) return ValidationResult.Ok();
                        if (!XqlCommon.TryToDouble(value!, out _))
                            return ValidationResult.Fail(ErrCode.E_TYPE_MISMATCH, "Expect REAL.");
                        break;
                    }
                case ColumnKind.Bool:
                    {
                        if (XqlCommon.IsNullish(value)) return ValidationResult.Ok();
                        if (!XqlCommon.TryToBool(value!, out _))
                            return ValidationResult.Fail(ErrCode.E_TYPE_MISMATCH, "Expect BOOL.");
                        break;
                    }
                case ColumnKind.Text:
                    {
                        if (XqlCommon.IsNullish(value))
                            return ValidationResult.Ok();
                        break;
                    }
                case ColumnKind.Json:
                    {
                        if (XqlCommon.IsNullish(value)) return ValidationResult.Ok();
                        var s = XqlCommon.NormalizeToString(value!);
                        try { _ = JToken.Parse(s); }
                        catch (Exception ex)
                        {
                            return ValidationResult.Fail(ErrCode.E_JSON_PARSE, $"JSON parse error: {ex.Message}");
                        }
                        break;
                    }
                case ColumnKind.Date:
                    {
                        if (XqlCommon.IsNullish(value)) return ValidationResult.Ok();
                        if (!XqlCommon.TryToDate(value!, out _))
                            return ValidationResult.Fail(ErrCode.E_TYPE_MISMATCH, "Expect DATE.");
                        break;
                    }
                default:
                    return ValidationResult.Fail(ErrCode.E_UNSUPPORTED, $"Unsupported type: {ct.Kind}");
            }

            return ValidationResult.Ok();
        }

        internal static string HeaderNameOf(string sheetName) => $"_XQL_HDR_{sheetName}";

        internal static bool TryGetHeaderMarker(Excel.Worksheet ws, out Excel.Range headerRange)
        {
            headerRange = null!;
            try
            {
                var key = HeaderNameOf(ws.Name);

                using var wb = SmartCom<Workbook>.Wrap((Excel.Workbook)ws.Parent);
                using var wbn = SmartCom<Names>.Wrap(wb.Value!.Names);
                using var wsn = SmartCom<Names>.Wrap(ws.Names);

                using var nmWB = SmartCom<Name>.Wrap(TryFindName(wbn.Value!, key));
                using var nmWS = SmartCom<Name>.Wrap(nmWB.Value == null ? TryFindName(wsn.Value!, key) : null);

                var nm = nmWB.Value ?? nmWS.Value;
                if (nm == null) return false;

                Excel.Range? rg = null;
                try { rg = nm.RefersToRange; } catch { rg = null; }
                if (rg == null) return false;

                headerRange = SmartCom<Range>.Wrap(rg).Detach()!;
                return true;
            }
            catch { return false; }
        }

        internal static void SetHeaderMarker(Excel.Worksheet ws, Excel.Range header)
        {
            var key = HeaderNameOf(ws.Name);

            using var wb = SmartCom<Workbook>.Wrap((Excel.Workbook)ws.Parent);
            using var wbn = SmartCom<Names>.Wrap(wb.Value!.Names);
            using var wsn = SmartCom<Names>.Wrap(ws.Names);

            // 기존 것(양쪽 범위) 안전 삭제
            TryDeleteName(wsn.Value!, key);
            TryDeleteName(wbn.Value!, key);

            // RefersTo: ='<시트명>'!$A$1:$D$1
            var sheetName = ws.Name.Replace("'", "''");
            var addr = header.Address[true, true, Excel.XlReferenceStyle.xlA1, false];
            var refersTo = $"='{sheetName}'!{addr}";

            using var nm = SmartCom<Name>.Wrap(wbn.Value!.Add(Name: key, RefersTo: refersTo));
            try { nm.Value!.Visible = false; } catch { }
        }

        internal static void ClearHeaderMarker(Excel.Worksheet ws)
        {
            try
            {
                var key = HeaderNameOf(ws.Name);
                using var wb = SmartCom<Workbook>.Wrap((Excel.Workbook)ws.Parent);
                using var wbn = SmartCom<Names>.Wrap(wb.Value!.Names);
                using var wsn = SmartCom<Names>.Wrap(ws.Names);

                TryDeleteName(wsn.Value!, key);
                TryDeleteName(wbn.Value!, key);
            }
            catch { /* ignore */ }
        }

        internal static bool IsSameRange(Excel.Range a, Excel.Range b)
            => a.Worksheet.Name == b.Worksheet.Name && a.Row == b.Row && a.Column == b.Column
               && a.Rows.Count == b.Rows.Count && a.Columns.Count == b.Columns.Count;

        private static string Escape(string s) => s.Replace("\\", "\\\\").Replace(":", "\\:");
        private static string Unescape(string s) => s.Replace("\\:", ":").Replace("\\\\", "\\");
        private static int ParseIntOrDefault(string s, int startIndex, int def = 0)
        {
            if (string.IsNullOrEmpty(s) || startIndex >= s.Length) return def;
            if (int.TryParse(s.Substring(startIndex), System.Globalization.NumberStyles.Integer,
                             System.Globalization.CultureInfo.InvariantCulture, out var v)) return v;
            return def;
        }

        private static Excel.Name? TryFindName(Excel.Names names, string key)
        {
            foreach (Excel.Name n in names)
            {
                try
                {
                    var nm = n.Name ?? string.Empty;
                    var nml = n.NameLocal ?? string.Empty;
                    if (nm.Equals(key, StringComparison.OrdinalIgnoreCase) ||
                        nml.Equals(key, StringComparison.OrdinalIgnoreCase))
                    {
                        return n; // 반환용
                    }
                }
                catch { /* skip */ }
            }
            return null;
        }

        private static void TryDeleteName(Excel.Names names, string key)
        {
            var nm = TryFindName(names, key);
            try { nm?.Delete(); } catch { /* ignore */ }
        }

        // ───────────────────────── 키 열 검색
        internal static int FindKeyColumnAbsolute(Excel.Range header, string keyName)
        {
            // 1) 지정 키명 우선
            if (!string.IsNullOrWhiteSpace(keyName))
            {
                int hit = FindBy(header, keyName);
                if (hit > 0) return hit;
            }
            // 2) 'id' → 'key' → 3) 첫 열
            int id = FindBy(header, "id");
            if (id > 0) return id;
            int key = FindBy(header, "key");
            if (key > 0) return key;
            return header.Column;

            static int FindBy(Excel.Range hdr, string name)
            {
                for (int i = 1; i <= hdr.Columns.Count; i++)
                {
                    using var c = SmartCom<Range>.Wrap(hdr.Cells[1, i]);
                    var n = (c.Value!.Value2 as string)?.Trim();
                    if (!string.IsNullOrEmpty(n) &&
                        string.Equals(n, name, StringComparison.OrdinalIgnoreCase))
                        return hdr.Column + i - 1; // 절대 열번호
                }
                return 0;
            }
        }

        // ── 상태 시트 보장 & 찾기
        internal static Excel.Worksheet EnsureStateSheet(Excel.Workbook wb)
        {
            foreach (Excel.Worksheet ws in wb.Worksheets)
            {
                if (ws.Name == StateSheetName) return ws;
            }
            var sh = (Excel.Worksheet)wb.Worksheets.Add();
            sh.Name = StateSheetName;
            sh.Visible = Excel.XlSheetVisibility.xlSheetVeryHidden;
            using var hdr = SmartCom<Range>.Wrap(sh.Range["A1", "B1"]);
            hdr.Value!.Value2 = new object[,] { { "key", "value" } };
            return sh;
        }

        // K/V 전체 읽기 (2행부터 끝까지)
        internal static Dictionary<string, string> StateReadAll(Excel.Workbook wb)
        {
            var map = new Dictionary<string, string>(StringComparer.Ordinal);
            var sh = EnsureStateSheet(wb);

            TryGetUsedBounds(sh, out var fr, out var fc, out var lr, out var lc);
            int last = lr;

            for (int r = 2; r <= last; r++)
            {
                using var kCell = SmartCom<Range>.Wrap(sh.Cells[r, 1] as Excel.Range);
                using var vCell = SmartCom<Range>.Wrap(sh.Cells[r, 2] as Excel.Range);
                var k = Convert.ToString(kCell.Value?.Value2) ?? "";
                var v = Convert.ToString(vCell.Value?.Value2) ?? "";
                if (!string.IsNullOrWhiteSpace(k))
                    map[k] = v;
            }
            return map;
        }

        // 여러 K/V 한 번에 upsert
        internal static void StateSetMany(Excel.Workbook wb, IReadOnlyDictionary<string, string> kv)
        {
            if (kv == null || kv.Count == 0) return;
            var sh = EnsureStateSheet(wb);

            var exist = StateReadAll(wb);

            TryGetUsedBounds(sh, out var fr, out var fc, out var lr, out var lc);
            int lastRow = lr <= 1 ? 1 : lr;
            int appendAt = Math.Max(2, lastRow + 1);

            var batch = new List<(int row, string key, string val)>();

            foreach (var p in kv)
            {
                if (exist.ContainsKey(p.Key))
                {
                    int hitRow = -1;
                    try
                    {
                        using var a2 = SmartCom<Range>.Wrap(sh.Cells[sh.Rows.Count, 1] as Excel.Range);
                        using var rg = SmartCom<Range>.Wrap(sh.Range["A2", a2.Value]);
                        using var hit = SmartCom<Range>.Wrap(rg.Value!.Find(What: p.Key,
                                                                 LookIn: Excel.XlFindLookIn.xlValues,
                                                                 LookAt: Excel.XlLookAt.xlWhole,
                                                                 SearchDirection: Excel.XlSearchDirection.xlNext));
                        if (hit.Value != null) hitRow = hit.Value.Row;
                    }
                    catch { }
                    batch.Add((hitRow > 0 ? hitRow : appendAt++, p.Key, p.Value ?? ""));
                }
                else
                {
                    batch.Add((appendAt++, p.Key, p.Value ?? ""));
                }
            }

            foreach (var b in batch)
            {
                using var c1 = SmartCom<Range>.Wrap(sh.Cells[b.row, 1] as Excel.Range);
                using var c2 = SmartCom<Range>.Wrap(sh.Cells[b.row, 2] as Excel.Range);
                if (c1.Value != null) c1.Value.Value2 = b.key;
                if (c2.Value != null) c2.Value.Value2 = b.val;
            }
        }

        // ───────────────────────── Shadow sheet (fingerprints)
        private static Excel.Worksheet EnsureShadowSheet(Excel.Workbook wb)
        {
            foreach (Excel.Worksheet ws in wb.Worksheets)
            {
                if (ws.Name == ShadowSheetName) return ws;
            }
            var sh = (Excel.Worksheet)wb.Worksheets.Add();
            sh.Name = ShadowSheetName;
            sh.Visible = Excel.XlSheetVisibility.xlSheetVeryHidden;
            using var hdr = SmartCom<Range>.Wrap(sh.Range["A1", "E1"]);
            hdr.Value!.Value2 = new object[,] { { "table", "row_key", "col_uid", "fp", "updated_utc" } };
            return sh;
        }

        internal static void ShadowAppendFingerprints(Excel.Workbook wb, IReadOnlyList<(string table, string rowKey, string colUid, string fp)> items)
        {
            if (items == null || items.Count == 0) return;
            var sh = EnsureShadowSheet(wb);

            TryGetUsedBounds(sh, out var fr, out var fc, out var lr, out var lc);
            int lastRow = lr <= 1 ? 1 : lr;

            int start = Math.Max(2, lastRow + 1);
            var data = new object[items.Count, 5];
            var now = DateTime.UtcNow.ToString("o");
            for (int i = 0; i < items.Count; i++)
            {
                data[i, 0] = items[i].table;
                data[i, 1] = items[i].rowKey;
                data[i, 2] = items[i].colUid;
                data[i, 3] = items[i].fp;
                data[i, 4] = now;
            }
            using var r1 = SmartCom<Range>.Wrap(sh.Cells[start, 1] as Range);
            using var r2 = SmartCom<Range>.Wrap(sh.Cells[start + items.Count - 1, 5] as Range);
            using var rg = SmartCom<Range>.Wrap(sh.Range[r1.Value, r2.Value]);
            rg.Value!.Value2 = data;
        }

        // (옵션) 헤더에서 col_uid를 얻는 간단 맵 — 최소 구현: 이름 자체를 uid로 사용
        internal static Dictionary<string, string> BuildColUidMap(Excel.Worksheet ws, Excel.Range header)
        {
            var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            for (int i = 1; i <= header.Columns.Count; i++)
            {
                using var hc = SmartCom<Range>.Wrap(header.Cells[1, i] as Excel.Range);
                var name = (string?)(hc.Value as Range)!.Value2 ?? "";
                if (!string.IsNullOrWhiteSpace(name))
                    map[name] = name;
            }
            return map;
        }
    }
}
