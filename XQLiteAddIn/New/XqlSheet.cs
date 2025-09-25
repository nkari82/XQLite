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

        /// <summary>툴팁 딕셔너리 생성.</summary>
        public Dictionary<string, string> BuildTooltipsForSheet(string sheetName)
        {
            if (!_sheets.TryGetValue(sheetName, out var sm)) return new();
            var dict = new Dictionary<string, string>(StringComparer.Ordinal);
            foreach (var kv in sm.Columns)
                dict[kv.Key] = kv.Value.ToTooltip();
            return dict;
        }

        /// <summary>
        /// 1행 헤더 Range와 트림된 헤더 텍스트 리스트를 반환.
        /// 반환된 header는 호출자가 ReleaseCom 해야 함(여기서 해제 금지).
        /// </summary>
        internal static (Excel.Range header, List<string> names) GetHeaderAndNames(Excel.Worksheet ws)
        {
            Excel.Range header = ws.Range[ws.Cells[1, 1], ws.Cells[1, ws.UsedRange.Columns.Count]];
            var names = new List<string>();
            int cols = header.Columns.Count;
            for (int c = 1; c <= cols; c++)
            {
                var cell = (Excel.Range)header.Cells[1, c];
                names.Add(Convert.ToString(cell.Value2) ?? string.Empty);
                XqlCommon.ReleaseCom(cell);
            }
            return (header, names.Select(s => s.Trim()).ToList());
        }

        // ───────────────────────── Header (fallback to UsedRange row1)
        public Excel.Range GetHeaderRange(Excel.Worksheet sh)
        {
            Excel.Range? lastCell = null;
            try
            {
                int col = 1;
                for (; col <= sh.UsedRange.Columns.Count; col++)
                {
                    var cell = (Excel.Range)sh.Cells[1, col];
                    var txt = (cell?.Value2 as string)?.Trim();
                    if (string.IsNullOrEmpty(txt)) { XqlCommon.ReleaseCom(cell); break; }
                    XqlCommon.ReleaseCom(lastCell); lastCell = cell;
                }
                var lastCol = Math.Max(1, (lastCell?.Column as int?) ?? 1);
                var rg = sh.Range[sh.Cells[1, 1], sh.Cells[1, lastCol]];
                XqlCommon.ReleaseCom(lastCell);
                return rg;
            }
            catch
            {
                XqlCommon.ReleaseCom(lastCell);
                return (Excel.Range)sh.Cells[1, 1];
            }
        }

        // ───────────────────────── Finders (인스턴스)
        internal static Excel.Worksheet? FindWorksheet(Excel.Application app, string sheetName)
        {
            if (app is null || string.IsNullOrWhiteSpace(sheetName)) return null;
            foreach (Excel.Worksheet? ws in app.Worksheets)
            {
                try
                {
                    if (ws is null) continue;
                    if (string.Equals(ws.Name, sheetName, StringComparison.Ordinal))
                        return ws;
                }
                finally { XqlCommon.ReleaseCom(ws); }
            }
            return null;
        }

        internal static Excel.ListObject? FindListObject(Excel.Worksheet ws, string listObjectName)
        {
            if (ws is null || string.IsNullOrWhiteSpace(listObjectName)) return null;
            foreach (Excel.ListObject? lo in ws.ListObjects)
            {
                try
                {
                    if (lo is null) continue;
                    if (string.Equals(lo.Name, listObjectName, StringComparison.Ordinal))
                        return lo;
                }
                finally { XqlCommon.ReleaseCom(lo); }
            }
            return null;
        }

        internal static Excel.ListObject? FindListObjectContaining(Excel.Worksheet ws, Excel.Range rng)
        {
            if (ws is null || rng is null) return null;
            foreach (Excel.ListObject? lo in ws.ListObjects)
            {
                Excel.Range? header = null;
                Excel.Range? body = null;
                try
                {
                    if (lo is null) continue;
                    header = lo.HeaderRowRange;
                    body = lo.DataBodyRange;

                    if ((header != null && XqlCommon.IntersectSafe(ws, rng, header) != null) ||
                        (body != null && XqlCommon.IntersectSafe(ws, rng, body) != null))
                        return lo;
                }
                catch { /* ignore */ }
                finally
                {
                    XqlCommon.ReleaseCom(header);
                    XqlCommon.ReleaseCom(body);
                    XqlCommon.ReleaseCom(lo);
                }
            }
            return null;
        }

        internal static Excel.ListObject? FindListObjectByTable(Excel.Worksheet ws, string tableNameOrQualified)
        {
            if (ws is null || string.IsNullOrWhiteSpace(tableNameOrQualified)) return null;

            // "Sheet.Table" 지원
            string? sheetHint = null;
            string loName = tableNameOrQualified;
            var dot = tableNameOrQualified.IndexOf('.');
            if (dot >= 0)
            {
                sheetHint = tableNameOrQualified.Substring(0, dot);
                loName = tableNameOrQualified.Substring(dot + 1);
            }

            // 1) 현재 시트 이름
            var found = FindListObject(ws, loName);
            if (found != null) return found;

            // 2) 헤더 캡션 추정
            found = FindListObjectByHeaderCaption(ws, loName);
            if (found != null) return found;

            // 3) sheetHint로 넘어가서 재시도
            if (!string.IsNullOrEmpty(sheetHint))
            {
                Excel.Worksheet? ws2 = null;
                try
                {
                    var app = ws.Application as Excel.Application;
                    ws2 = FindWorksheet(app!, sheetHint);
                    if (ws2 != null)
                    {
                        var f2 = FindListObject(ws2, loName) ?? FindListObjectByHeaderCaption(ws2, loName);
                        if (f2 != null) return f2;
                    }
                }
                finally { XqlCommon.ReleaseCom(ws2); }
            }

            // 4) 정규화 "Sheet.Table" 최종 시도
            foreach (Excel.ListObject? lo in ws.ListObjects)
            {
                try
                {
                    if (lo is null) continue;
                    var qualified = $"{ws.Name}.{lo.Name}";
                    if (string.Equals(qualified, tableNameOrQualified, StringComparison.Ordinal))
                        return lo;
                }
                finally { XqlCommon.ReleaseCom(lo); }
            }

            return null;
        }

        private static Excel.ListObject? FindListObjectByHeaderCaption(Excel.Worksheet ws, string caption)
        {
            foreach (Excel.ListObject? lo in ws.ListObjects)
            {
                Excel.Range? header = null;
                try
                {
                    if (lo is null) continue;
                    header = lo.HeaderRowRange;
                    if (header == null) continue;

                    var v = header.Value2 as object[,];
                    if (v == null) continue;

                    string first = Convert.ToString(v[1, 1]) ?? string.Empty;
                    if (string.Equals(first, caption, StringComparison.Ordinal))
                        return lo;

                    int cols = header.Columns.Count;
                    var joined = string.Empty;
                    for (int i = 1; i <= cols; i++)
                        joined += (i == 1 ? "" : "|") + (Convert.ToString(v[1, i]) ?? string.Empty);
                    if (string.Equals(joined, caption, StringComparison.Ordinal))
                        return lo;
                }
                catch { /* ignore */ }
                finally { XqlCommon.ReleaseCom(header); XqlCommon.ReleaseCom(lo); }
            }
            return null;
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

                lo = FindListObject(ws, d.Table);
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

        // ───────────────────────── Legacy Lock Key Migration ─────────────────────────
        private static readonly Regex RxOldCell =
            new(@"^cell:(?<sheet>[^!]+)!(?<addr>\$?[A-Z]+\$?\d+)$",
                RegexOptions.Compiled | RegexOptions.CultureInvariant | RegexOptions.IgnoreCase);

        private static readonly Regex RxOldColumn =
            new(@"^(?:column|col):(?<table>[^\.]+)\.(?<column>.+)$",
                RegexOptions.Compiled | RegexOptions.CultureInvariant | RegexOptions.IgnoreCase);

        /// <summary>
        /// 구 포맷 키(cell:Sheet!A1, column:Table.Col)를 상대키(col:/cell:)로 변환.
        /// 인식 불가능하면 입력을 그대로 반환.
        /// </summary>
        internal static string MigrateLockKeyIfNeeded(Excel.Application app, string oldKey)
        {
            if (string.IsNullOrWhiteSpace(oldKey)) return oldKey;

            // 1) cell:Sheet!A1
            var mCell = RxOldCell.Match(oldKey);
            if (mCell.Success)
            {
                var sheetName = mCell.Groups["sheet"].Value;
                var addr = mCell.Groups["addr"].Value;

                var ws = FindWorksheet(app, sheetName);
                if (ws == null) return oldKey;

                Excel.Range? rng = null; Excel.ListObject? lo = null;
                try
                {
                    rng = ws.Range[addr];
                    lo = rng?.ListObject ?? FindListObjectContaining(ws, rng!);
                    if (lo?.HeaderRowRange == null) return oldKey;

                    int hRow = lo.HeaderRowRange.Row;
                    int hCol = lo.HeaderRowRange.Column;

                    int rowOffset = rng!.Row - (hRow + 1); // 데이터 첫행 = header+1
                    int colOffset = rng!.Column - hCol;

                    var tableName = XqlTableNameMap.Map(lo.Name, ws.Name);
                    return CellKey(ws.Name, tableName, hRow, hCol, rowOffset, colOffset);
                }
                catch { return oldKey; }
                finally { XqlCommon.ReleaseCom(lo); XqlCommon.ReleaseCom(rng); }
            }

            // 2) column:Table.Column
            var mCol = RxOldColumn.Match(oldKey);
            if (mCol.Success)
            {
                var table = mCol.Groups["table"].Value;
                var col = mCol.Groups["column"].Value;

                Excel.Worksheet? ws = null; Excel.ListObject? lo = null;
                try
                {
                    ws = app.ActiveSheet as Excel.Worksheet;
                    if (ws == null) return oldKey;

                    lo = FindListObjectByTable(ws, table);
                    if (lo?.HeaderRowRange == null) return oldKey;

                    // 표 헤더에서 정확히 컬럼 인덱스 매칭(0-base)
                    int colIndex = -1;
                    var v = lo.HeaderRowRange.Value2 as object[,];
                    if (v != null)
                    {
                        int cols = lo.HeaderRowRange.Columns.Count;
                        for (int c = 1; c <= cols; c++)
                        {
                            var name = (Convert.ToString(v[1, c]) ?? string.Empty).Trim();
                            if (string.Equals(name, col, StringComparison.Ordinal))
                            { colIndex = c - 1; break; }
                        }
                    }
                    if (colIndex < 0) return oldKey;

                    int hRow = lo.HeaderRowRange.Row;
                    int hCol = lo.HeaderRowRange.Column;

                    return ColumnKey(ws.Name, table, hRow, hCol, colIndex, col);
                }
                catch { return oldKey; }
                finally { XqlCommon.ReleaseCom(lo); XqlCommon.ReleaseCom(ws); }
            }

            // 3) 이미 신 포맷/미인식 포맷
            return oldKey;
        }

        // ───────────────────────── 기타 유틸
        internal static int? FindRowByKey(Excel.Worksheet ws, int firstDataRow, int keyCol, object key)
        {
            try
            {
                var used = ws.UsedRange;
                int lastRow = used.Row + used.Rows.Count - 1;
                XqlCommon.ReleaseCom(used);
                for (int r = firstDataRow; r <= lastRow; r++)
                {
                    Excel.Range? cell = null;
                    try
                    {
                        cell = (Excel.Range)ws.Cells[r, keyCol];
                        var v = cell.Value2;
                        if (XqlCommon.EqualKey(v, key))
                            return r;
                    }
                    finally
                    {
                        XqlCommon.ReleaseCom(cell);
                    }
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
            if (!_sheets.TryGetValue(sheet, out var sm)) return ValidationResult.Ok();
            if (!sm.Columns.TryGetValue(col, out var ct)) return ValidationResult.Ok();

            // NotNull
            if (!ct.Nullable && XqlCommon.IsNullish(value))
                return ValidationResult.Fail(ErrCode.E_NULL_NOT_ALLOWED, "Null/empty is not allowed.");

            // 타입별 검증
            switch (ct.Kind)
            {
                case ColumnKind.Int:
                    {
                        if (XqlCommon.IsNullish(value)) return ValidationResult.Ok();
                        if (!XqlCommon.TryToInt64(value, out var iv))
                            return ValidationResult.Fail(ErrCode.E_TYPE_MISMATCH, "Expect INT.");
                        if (ct.Min.HasValue && iv < (long)ct.Min.Value)
                            return ValidationResult.Fail(ErrCode.E_RANGE, $"INT < Min({ct.Min})");
                        if (ct.Max.HasValue && iv > (long)ct.Max.Value)
                            return ValidationResult.Fail(ErrCode.E_RANGE, $"INT > Max({ct.Max})");
                        break;
                    }
                case ColumnKind.Real:
                    {
                        if (XqlCommon.IsNullish(value)) return ValidationResult.Ok();
                        if (!XqlCommon.TryToDouble(value, out var dv))
                            return ValidationResult.Fail(ErrCode.E_TYPE_MISMATCH, "Expect REAL.");
                        if (ct.Min.HasValue && dv < ct.Min.Value)
                            return ValidationResult.Fail(ErrCode.E_RANGE, $"REAL < Min({ct.Min})");
                        if (ct.Max.HasValue && dv > ct.Max.Value)
                            return ValidationResult.Fail(ErrCode.E_RANGE, $"REAL > Max({ct.Max})");
                        break;
                    }
                case ColumnKind.Bool:
                    {
                        if (XqlCommon.IsNullish(value)) return ValidationResult.Ok();
                        if (!XqlCommon.TryToBool(value, out _))
                            return ValidationResult.Fail(ErrCode.E_TYPE_MISMATCH, "Expect BOOL.");
                        break;
                    }
                case ColumnKind.Text:
                    {
                        if (XqlCommon.IsNullish(value)) return ValidationResult.Ok();
                        var s = XqlCommon.NormalizeToString(value);
                        if (ct.Regex != null && !ct.Regex.IsMatch(s))
                            return ValidationResult.Fail(ErrCode.E_CHECK_FAIL, "TEXT regex mismatch.");
                        break;
                    }
                case ColumnKind.Json:
                    {
                        if (XqlCommon.IsNullish(value)) return ValidationResult.Ok();
                        var s = XqlCommon.NormalizeToString(value);
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
                        if (!XqlCommon.TryToDate(value, out _))
                            return ValidationResult.Fail(ErrCode.E_TYPE_MISMATCH, "Expect DATE.");
                        break;
                    }
                default:
                    return ValidationResult.Fail(ErrCode.E_UNSUPPORTED, $"Unsupported type: {ct.Kind}");
            }

            // 사용자 정의 체크
            if (ct.CustomCheck != null)
            {
                try
                {
                    if (!ct.CustomCheck(value))
                        return ValidationResult.Fail(ErrCode.E_CHECK_FAIL, ct.CustomCheckDescription ?? "Custom check failed.");
                }
                catch (Exception ex)
                {
                    return ValidationResult.Fail(ErrCode.E_CHECK_FAIL, $"Custom check error: {ex.Message}");
                }
            }

            return ValidationResult.Ok();
        }


        private static string Escape(string s) => s.Replace("\\", "\\\\").Replace(":", "\\:");
        private static string Unescape(string s) => s.Replace("\\:", ":").Replace("\\\\", "\\");
        private static int ParseIntOrDefault(string s, int startIndex, int def = 0)
        {
            if (string.IsNullOrEmpty(s) || startIndex >= s.Length) return def;
            if (int.TryParse(s.Substring(startIndex), System.Globalization.NumberStyles.Integer,
                             System.Globalization.CultureInfo.InvariantCulture, out var v)) return v;
            return def;
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

    internal enum ColumnKind { Text, Int, Real, Bool, Date, Json }

    internal sealed class ColumnType
    {
        public ColumnKind Kind { get; set; } = ColumnKind.Text;
        public bool Nullable { get; set; } = true;

        /// <summary>수치형에만 사용(Int=long, Real=double)</summary>
        public double? Min;
        public double? Max;

        /// <summary>Text 전용</summary>
        public Regex? Regex;

        /// <summary>사용자 정의 체크(선택)</summary>
        public Func<object?, bool>? CustomCheck = null;
        public string? CustomCheckDescription = null;

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
