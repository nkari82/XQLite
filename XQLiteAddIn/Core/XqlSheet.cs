// XqlSheet.cs
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;

namespace XQLite.AddIn
{
    /// <summary>
    /// 시트 범용 유틸(찾기/상대키/헤더 Fallback 등) + "메타 레지스트리" 기능을 흡수한 인스턴스 서비스.
    /// UI 로직은 XqlSheetView가 담당한다.
    /// </summary>
    internal sealed class XqlSheet
    {
        private readonly Dictionary<string, SheetMeta> _sheets = new(StringComparer.Ordinal);

        // ───────────────────────── Meta registry (흡수)
        public bool TryGetSheet(string sheetName, out SheetMeta sm) => _sheets.TryGetValue(sheetName, out sm!);

        public SheetMeta GetOrCreateSheet(string sheetName, string defaultKeyColumn = "id")
        {
            if (!_sheets.TryGetValue(sheetName, out var sm))
            {
                sm = new SheetMeta { TableName = sheetName, KeyColumn = defaultKeyColumn };
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

                if (!sm.Columns.TryGetValue(name, out var ct))
                {
                    ct = new ColumnType { Kind = defaultKind, Nullable = defaultNullable };
                    sm.SetColumn(name, ct);
                    added.Add(name);
                }
            }
            return added;
        }

        /// <summary>
        /// 1행 헤더 Range와 트림된 헤더 텍스트 리스트를 반환.
        /// 반환된 header는 호출자가 ReleaseCom 해야 함(여기서 해제 금지).
        /// </summary>
        internal static (Excel.Range header, List<string> names) GetHeaderAndNames(Excel.Worksheet ws)
        {
            int firstCol = 1, lastCol = 1;
            Excel.Range? used = null, cell = null;
            try
            {
                try
                {
                    used = ws.UsedRange;
                    firstCol = used.Column;
                    lastCol = used.Column + used.Columns.Count - 1;
                }
                catch
                {
                    firstCol = 1; lastCol = 1;
                }
                finally
                {
                    XqlCommon.ReleaseCom(used);
                }

                // 경계 보정
                if (lastCol < firstCol) lastCol = firstCol;

                // 1행의 [firstCol..lastCol]을 헤더로 간주
                var header = ws.Range[ws.Cells[1, firstCol], ws.Cells[1, lastCol]];

                int colCount = header.Columns.Count;
                var names = new List<string>(capacity: colCount);

                for (int i = 1; i <= colCount; i++)
                {
                    try
                    {
                        // ⚠️ 헤더 내부 상대 인덱스 사용 (1-based)
                        cell = (Excel.Range)header.Cells[1, i];

                        var raw = (Convert.ToString(cell.Value2) ?? "").Trim();
                        // 빈 칸이면 A,B,C… 폴백
                        names.Add(string.IsNullOrEmpty(raw)
                                  ? XqlCommon.ColumnIndexToLetter(firstCol + i - 1)
                                  : raw!);
                    }
                    finally
                    {
                        XqlCommon.ReleaseCom(cell);
                        cell = null;
                    }
                }

                // header는 호출자가 ReleaseCom 해야 함
                return (header, names);
            }
            catch
            {
                // 최소 안전 폴백
                var header = ws.Range[ws.Cells[1, 1], ws.Cells[1, 1]];
                return (header, new List<string> { XqlCommon.ColumnIndexToLetter(1) });
            }
        }


        // ───────────────────────── Header (fallback to UsedRange row1)
        // XqlSheetView.cs에서만 사용하므로 옮길수 도 있다.
        internal static Excel.Range GetHeaderRange(Excel.Worksheet sh)
        {
            Excel.Range? used = null;
            try
            {
                used = sh.UsedRange;
                int lastCol = Math.Max(1, used.Column + used.Columns.Count - 1);
                int hr = 1; // 헤더는 1행 기준 (마커/ResolveHeader가 우선이며, 이는 Fallback)
                int last = 1;
                for (int c = lastCol; c >= 1; --c)
                {
                    Excel.Range? cell = null;
                    try
                    {
                        cell = (Excel.Range)sh.Cells[hr, c];
                        var txt = (Convert.ToString(cell.Value2) ?? "").Trim();
                        if (!string.IsNullOrEmpty(txt)) { last = c; break; }
                    }
                    finally { XqlCommon.ReleaseCom(cell); }
                }
                return sh.Range[sh.Cells[hr, 1], sh.Cells[hr, last]];
            }
            catch
            {
                return (Excel.Range)sh.Cells[1, 1];
            }
            finally { XqlCommon.ReleaseCom(used); }
        }

        // ───────────────────────── Finders (인스턴스)
        internal static Excel.Worksheet? FindWorksheet(Excel.Application app, string sheetName)
        {
            if (app is null || string.IsNullOrWhiteSpace(sheetName)) return null;

            Excel.Worksheet? match = null;
            foreach (Excel.Worksheet? ws in app.Worksheets)
            {
                if (ws is null) continue;
                if (string.Equals(ws.Name, sheetName, StringComparison.Ordinal))
                {
                    match = ws;       // ★ 선택한 객체는 해제하지 않음 (호출자가 책임)
                    break;
                }
                XqlCommon.ReleaseCom(ws); // 매칭 실패한 것만 해제
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
                    match = lo;       // ★ 선택한 객체는 해제하지 않음
                    break;
                }
                XqlCommon.ReleaseCom(lo);
            }
            return match;
        }

        internal static Excel.ListObject? FindListObjectContaining(Excel.Worksheet ws, Excel.Range rng)
        {
            if (ws is null || rng is null) return null;

            Excel.ListObject? match = null;

            foreach (Excel.ListObject lo in ws.ListObjects)
            {
                bool keep = false;
                Excel.Range? header = null, body = null;
                Excel.Range? inter1 = null, inter2 = null;

                try
                {
                    header = lo.HeaderRowRange;
                    body = lo.DataBodyRange;

                    bool hit = false;
                    if (header != null)
                    {
                        inter1 = XqlCommon.IntersectSafe(ws, rng, header);
                        hit |= inter1 != null;
                    }
                    if (!hit && body != null)
                    {
                        inter2 = XqlCommon.IntersectSafe(ws, rng, body);
                        hit |= inter2 != null;
                    }

                    if (hit)
                    {
                        match = lo;   // 매치된 객체는 호출자에게 반환(해제 금지)
                        keep = true;
                        break;
                    }
                }
                finally
                {
                    XqlCommon.ReleaseCom(inter2);
                    XqlCommon.ReleaseCom(inter1);
                    XqlCommon.ReleaseCom(body);
                    XqlCommon.ReleaseCom(header);
                    if (!keep) XqlCommon.ReleaseCom(lo); // 매치 실패건만 해제
                }
            }

            return match;
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
                Excel.Worksheet? ws2 = null;
                try
                {
                    ws2 = FindWorksheet(app, sheetHint!);
                    if (ws2 != null)
                    {
                        var f2 = FindListObject(ws2, loName) ?? FindListObjectByHeaderCaption(ws2, loName);
                        if (f2 != null) return f2;
                    }
                }
                finally { XqlCommon.ReleaseCom(ws2); }
            }

            Excel.ListObject? match = null;
            foreach (Excel.ListObject? lo in ws.ListObjects)
            {
                if (lo is null) continue;
                var qualified = $"{ws.Name}.{lo.Name}";
                if (string.Equals(qualified, tableNameOrQualified, StringComparison.Ordinal))
                {
                    match = lo;              // ← 매치된 것은 반환(해제 금지)
                    break;
                }
                XqlCommon.ReleaseCom(lo);    // 매치 실패건만 해제
            }

            // 4) 폴백: 통합문서 전체 스캔 (비싼 연산이므로 최후순위)
            try
            {
                var app2 = ws.Application as Excel.Application;
                if (app2 != null)
                {
                    foreach (Excel.Worksheet w in app2.Worksheets)
                    {
                        try
                        {
                            foreach (Excel.ListObject lo2 in w.ListObjects)
                            {
                                try
                                {
                                    if (string.Equals(lo2.Name, loName, StringComparison.Ordinal) ||
    string.Equals($"{w.Name}.{lo2.Name}", tableNameOrQualified, StringComparison.Ordinal))
                                        return lo2; // ← 매치된 것은 반환(해제 금지)
                                }
                                finally { if (!object.ReferenceEquals(match, lo2)) XqlCommon.ReleaseCom(lo2); }
                            }
                        }
                        finally { if (!object.ReferenceEquals(match, w)) XqlCommon.ReleaseCom(w); }
                    }
                }
            }
            catch { }
            return null;
        }

        private static Excel.ListObject? FindListObjectByHeaderCaption(Excel.Worksheet ws, string caption)
        {
            Excel.ListObject? match = null;

            foreach (Excel.ListObject? lo in ws.ListObjects)
            {
                if (lo is null) continue;
                Excel.Range? header = null;
                try
                {
                    header = lo.HeaderRowRange;
                    if (header == null) continue;

                    var v = header.Value2 as object[,];
                    if (v == null) continue;

                    string first = Convert.ToString(v[1, 1]) ?? string.Empty;
                    if (string.Equals(first, caption, StringComparison.Ordinal))
                    {
                        match = lo;                 // ★ 선택된 객체는 해제 금지
                        break;
                    }

                    int cols = header.Columns.Count;
                    var joined = string.Empty;
                    for (int i = 1; i <= cols; i++)
                        joined += (i == 1 ? "" : "|") + (Convert.ToString(v[1, i]) ?? string.Empty);
                    if (string.Equals(joined, caption, StringComparison.Ordinal))
                    {
                        match = lo;
                        break;
                    }
                }
                finally
                {
                    XqlCommon.ReleaseCom(header);
                    if (!object.ReferenceEquals(match, lo))
                        XqlCommon.ReleaseCom(lo);
                }
            }
            return match;
        }

        internal static int FindKeyColumnIndex(List<string> headers, string keyName) { if (!string.IsNullOrWhiteSpace(keyName)) { var idx = headers.FindIndex(h => string.Equals(h, keyName, StringComparison.Ordinal)); if (idx >= 0) return idx + 1; } var id = headers.FindIndex(h => string.Equals(h, "id", StringComparison.OrdinalIgnoreCase)); if (id >= 0) return id + 1; var key = headers.FindIndex(h => string.Equals(h, "key", StringComparison.OrdinalIgnoreCase)); if (key >= 0) return key + 1; return 1; }

        // ───────────────────────── Relative Keys (Serialize/Parse/Resolve) : 인스턴스 + 하위호환 static 포워딩

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

        internal static bool TryResolve(Excel.Application app, RelDesc d, out Excel.Range? target, out int? targetHeaderCol, out Excel.ListObject? lo)
        {
            target = null; targetHeaderCol = null; lo = null;
            try
            {
                var ws = FindWorksheet(app, d.Sheet);
                if (ws == null) return false;

                lo = FindListObject(ws, d.Table) ?? FindListObjectByTable(ws, d.Table);
                if (lo == null || lo.HeaderRowRange == null) return false;

                var hdr = lo.HeaderRowRange;
                int anchorRow = hdr.Row;
                int anchorCol = hdr.Column;

                if (d.Kind == "col")
                {
                    int col = anchorCol + d.ColOffset;
                    targetHeaderCol = col;
                    target = (Excel.Range)ws.Cells[anchorRow, col];
                    return true;
                }
                else if (d.Kind == "cell")
                {
                    int row = (anchorRow + 1) + d.RowOffset;   // 데이터 첫 행 = header+1
                    int col = anchorCol + d.ColOffset;
                    target = (Excel.Range)ws.Cells[row, col];
                    return true;
                }
                return false;
            }
            catch { return false; }
        }

        // ── Back-compat static forwards (기존 호출부 보호)
#if false
        public static Excel.Worksheet? FindWorksheet(Excel.Application app, string sheetName)
            => XqlAddIn.SheetSvc?.FindWorksheet(app, sheetName);

        public static Excel.ListObject? FindListObject(Excel.Worksheet ws, string listObjectName)
            => XqlAddIn.SheetSvc?.FindListObject(ws, listObjectName);

        public static Excel.ListObject? FindListObjectContaining(Excel.Worksheet ws, Excel.Range rng)
            => XqlAddIn.SheetSvc?.FindListObjectContaining(ws, rng);

        public static Excel.ListObject? FindListObjectByTable(Excel.Worksheet ws, string tableNameOrQualified)
            => XqlAddIn.SheetSvc?.FindListObjectByTable(ws, tableNameOrQualified);

        public static string ColumnKey(string sheet, string table, int hRow, int hCol, int colOffset, string? hdrName = null)
            => XqlAddIn.SheetSvc?.ColumnKey(sheet, table, hRow, hCol, colOffset, hdrName) ?? "";

        public static string CellKey(string sheet, string table, int hRow, int hCol, int rowOffset, int colOffset)
            => XqlAddIn.SheetSvc?.CellKey(sheet, table, hRow, hCol, rowOffset, colOffset) ?? "";

        public static bool TryParse(string key, out RelDesc d)
            => XqlAddIn.SheetSvc?.TryParse(key, out d) ?? (d = default, false).Item2;

        internal static bool TryResolve(Excel.Application app, RelDesc d, out Excel.Range? target, out int? targetHeaderCol, out Excel.ListObject? lo)
            => XqlAddIn.SheetSvc?.TryResolve(app, d, out target, out targetHeaderCol, out lo) ?? (target = null, targetHeaderCol = null, lo = null, false).Item4;
#endif
        // ───────────────────────── 기타 유틸
        internal static int? FindRowByKey(Excel.Worksheet ws, int firstDataRow, int keyCol, object key)
        {
            try
            {
                var used = ws.UsedRange;
                int lastRow = used.Row + used.Rows.Count - 1;
                XqlCommon.ReleaseCom(used);

                // 1) 빠른 경로: Range.Find (정확 일치)
                Excel.Range? rg = null, hit = null;
                try
                {
                    rg = ws.Range[ws.Cells[firstDataRow, keyCol], ws.Cells[lastRow, keyCol]];
                    hit = rg.Find(What: key, After: Type.Missing,
                                  LookIn: Excel.XlFindLookIn.xlValues,
                                  LookAt: Excel.XlLookAt.xlWhole,
                                  SearchOrder: Excel.XlSearchOrder.xlByRows,
                                  SearchDirection: Excel.XlSearchDirection.xlNext,
                                  MatchCase: false);
                    if (hit != null) return hit.Row;
                }
                catch { /* fall through */ }
                finally { XqlCommon.ReleaseCom(hit); XqlCommon.ReleaseCom(rg); }

                // 2) 폴백: 선형 탐색 (기존 동작)
                for (int r = firstDataRow; r <= lastRow; r++)
                {
                    Excel.Range? cell = null;
                    try
                    {
                        cell = (Excel.Range)ws.Cells[r, keyCol];
                        var v = cell.Value2;
                        if (XqlCommon.EqualKey(v, key)) return r;
                    }
                    finally { XqlCommon.ReleaseCom(cell); }
                }
            }
            catch { }
            return null;
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
                        if (!XqlCommon.TryToInt64(value!, out var iv))
                            return ValidationResult.Fail(ErrCode.E_TYPE_MISMATCH, "Expect INT.");
                        break;
                    }
                case ColumnKind.Real:
                    {
                        if (XqlCommon.IsNullish(value)) return ValidationResult.Ok();
                        if (!XqlCommon.TryToDouble(value!, out var dv))
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
            Excel.Workbook? wb = null;
            Excel.Names? wbNames = null, wsNames = null;
            Excel.Name? nm = null;
            try
            {
                var key = HeaderNameOf(ws.Name);

                // 1) Workbook 범위 먼저
                wb = (Excel.Workbook)ws.Parent;
                wbNames = wb.Names;
                nm = TryFindName(wbNames, key);

                // 2) 없으면 Worksheet 범위
                if (nm == null)
                {
                    wsNames = ws.Names;
                    nm = TryFindName(wsNames, key);
                }

                if (nm == null) return false;

                Excel.Range? rg = null;
                try { rg = nm.RefersToRange; }     // 일부 이름은 RefersToRange 접근 시 예외 → null 처리
                catch { rg = null; }
                if (rg == null) return false;

                headerRange = rg;                   // 반환: 호출자가 ReleaseCom 해야 함
                return true;
            }
            catch { return false; }
            finally
            {
                XqlCommon.ReleaseCom(nm);
                XqlCommon.ReleaseCom(wsNames);
                XqlCommon.ReleaseCom(wbNames);
                XqlCommon.ReleaseCom(wb);
            }
        }

        internal static void SetHeaderMarker(Excel.Worksheet ws, Excel.Range header)
        {
            Excel.Workbook? wb = null;
            Excel.Names? wbNames = null, wsNames = null;
            Excel.Name? nm = null;
            try
            {
                var key = HeaderNameOf(ws.Name);

                wb = (Excel.Workbook)ws.Parent;
                wbNames = wb.Names;
                wsNames = ws.Names;

                // 기존 것(양쪽 범위) 안전 삭제
                TryDeleteName(wsNames, key);
                TryDeleteName(wbNames, key);

                // RefersTo: ='<시트명>'!$A$1:$D$1  (시트명 홑따옴표 이스케이프)
                var sheetName = ws.Name.Replace("'", "''");
                var addr = header.Address[true, true, Excel.XlReferenceStyle.xlA1, false];
                var refersTo = $"='{sheetName}'!{addr}";

                nm = wbNames.Add(Name: key, RefersTo: refersTo);
                try { nm.Visible = false; } catch { }
            }
            finally
            {
                XqlCommon.ReleaseCom(nm);
                XqlCommon.ReleaseCom(wsNames);
                XqlCommon.ReleaseCom(wbNames);
                XqlCommon.ReleaseCom(wb);
            }
        }

        internal static void ClearHeaderMarker(Excel.Worksheet ws)
        {
            Excel.Workbook? wb = null;
            Excel.Names? wbNames = null, wsNames = null;
            try
            {
                var key = HeaderNameOf(ws.Name);
                wb = (Excel.Workbook)ws.Parent;
                wbNames = wb.Names;
                wsNames = ws.Names;

                TryDeleteName(wsNames, key);
                TryDeleteName(wbNames, key);
            }
            catch { /* ignore */ }
            finally
            {
                XqlCommon.ReleaseCom(wsNames);
                XqlCommon.ReleaseCom(wbNames);
                XqlCommon.ReleaseCom(wb);
            }
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

        // XqlSheet.cs 내부(클래스 private static 영역)
        private static Excel.Name? TryFindName(Excel.Names names, string key)
        {
            Excel.Name? hit = null;
            foreach (Excel.Name n in names)
            {
                try
                {
                    var nm = n.Name ?? string.Empty;
                    var nml = n.NameLocal ?? string.Empty;
                    if (nm.Equals(key, StringComparison.OrdinalIgnoreCase) ||
                        nml.Equals(key, StringComparison.OrdinalIgnoreCase))
                    {
                        hit = n;                 // 매치: 즉시 종료
                        break;
                    }
                }
                finally
                {
                    if (!object.ReferenceEquals(hit, n)) XqlCommon.ReleaseCom(n); // 매치 실패건만 해제
                }
            }
            return hit;
        }

        private static void TryDeleteName(Excel.Names names, string key)
        {
            var nm = TryFindName(names, key);
            try { nm?.Delete(); }
            catch { /* ignore */ }
            finally { XqlCommon.ReleaseCom(nm); }
        }


        // XqlSheet.cs — 아래 유틸 추가 (파일 내 임의 위치: 다른 static 메서드 근처)
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
                    Excel.Range? c = null;
                    try
                    {
                        c = (Excel.Range)hdr.Cells[1, i];
                        var n = (c.Value2 as string)?.Trim();
                        if (!string.IsNullOrEmpty(n) &&
                            string.Equals(n, name, StringComparison.OrdinalIgnoreCase))
                            return hdr.Column + i - 1; // 절대 열번호
                    }
                    finally { XqlCommon.ReleaseCom(c); }
                }
                return 0;
            }
        }


    }

    // ───────────────────────── Models (같은 파일에 유지)
    internal sealed class SheetMeta
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
}
