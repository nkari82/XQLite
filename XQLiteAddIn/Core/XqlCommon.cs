// XqlCommon.cs
using ExcelDna.Integration;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;


namespace XQLite.AddIn
{
    internal static class XqlCommon
    {
        private static readonly object _sumLock = new();
        private static HashSet<string> _sumTables = new(StringComparer.Ordinal);
        private static int _sumAffected, _sumConflicts, _sumErrors, _sumBatches;
        private static long _sumStartTicks;


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
                    Put(ws, 1, 1, title, bold: true, size: 16);
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
                    finally { ReleaseCom(box); }

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
                        finally { ReleaseCom(cell); }
                    }
                }
                catch { }
                finally { ReleaseCom(r); ReleaseCom(ws); ReleaseCom(wb); }
            });

            static double TicksToMs(long ticks)
            {
                double freq = System.Diagnostics.Stopwatch.Frequency;
                return ticks * 1000.0 / freq;
            }
        }

        // Excel Column Index -> "A, B, ..., Z, AA ..." 폴백 헤더명
        internal static string ColumnIndexToLetter(int col)
        {
            string s = string.Empty;
            while (col > 0)
            {
                int m = (col - 1) % 26;
                s = (char)('A' + m) + s;
                col = (col - 1) / 26;
            }
            return s;
        }

        internal static void ReleaseCom(object? o)
        {
            try { if (o != null && Marshal.IsComObject(o)) Marshal.FinalReleaseComObject(o); } catch { }
        }

        internal static string CreateTempDir(string prefix)
        {
            var root = Path.Combine(Path.GetTempPath(), prefix + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(root);
            return root;
        }

        internal static void TryDeleteDir(string dir)
        {
            try
            {
                if (Directory.Exists(dir))
                    Directory.Delete(dir, true);
            }
            catch { }
        }

        internal static void SafeZipDirectory(string dir, string outZip)
        {
            try
            {
                if (File.Exists(outZip))
                    File.Delete(outZip);

                ZipFile.CreateFromDirectory(dir, outZip, CompressionLevel.Fastest, includeBaseDirectory: false);
            }
            catch { }
        }

        internal static string CsvEscape(string s)
        {
            if (s == null) return "";
            bool needQuote = s.Contains(',') || s.Contains('"') || s.Contains('\n') || s.Contains('\r');
            return needQuote ? "\"" + s.Replace("\"", "\"\"") + "\"" : s;
        }

        internal static string ValueToString(object? v)
        {
            if (v == null) return "";
            if (v is bool b) return b ? "TRUE" : "FALSE";
            if (v is DateTime dt) return dt.ToString("o");
            if (v is double d) return d.ToString("R", CultureInfo.InvariantCulture);
            if (v is float f) return ((double)f).ToString("R", CultureInfo.InvariantCulture);
            if (v is decimal m) return m.ToString(CultureInfo.InvariantCulture);
            return Convert.ToString(v, CultureInfo.InvariantCulture) ?? "";
        }

        internal static IEnumerable<List<T>> Chunk<T>(IReadOnlyList<T> list, int size)
        {
            if (size <= 0) size = 1000;
            for (int i = 0; i < list.Count; i += size)
                yield return list.Skip(i).Take(Math.Min(size, list.Count - i)).ToList();
        }

        internal static void InterlockedMax(ref long target, long value)
        {
            while (true)
            {
                long cur = System.Threading.Volatile.Read(ref target);
                if (value <= cur) return;
                if (System.Threading.Interlocked.CompareExchange(ref target, value, cur) == cur) return;
            }
        }

        internal static Excel.Range? IntersectSafe(Excel.Worksheet ws, Excel.Range a, Excel.Range b)
        {
            try { return ws.Application.Intersect(a, b); }
            catch { return null; }
        }

        internal static bool EqualKey(object? a, object? b)
        {
            if (a is null && b is null) return true;
            if (a is null || b is null) return false;
            var sa = (Convert.ToString(a, CultureInfo.InvariantCulture) ?? "").Trim();
            var sb = (Convert.ToString(b, CultureInfo.InvariantCulture) ?? "").Trim();
            return string.Equals(sa, sb, StringComparison.Ordinal);
        }

        internal static bool IsNullish(object? v)
        {
            if (v is null) return true;
            if (v is string s) return string.IsNullOrWhiteSpace(s);
            return false;
        }

        internal static bool TryToInt64(object v, out long value)
        {
            try
            {
                switch (v)
                {
                    case sbyte sb: value = sb; return true;
                    case byte b: value = b; return true;
                    case short s: value = s; return true;
                    case ushort us: value = us; return true;
                    case int i: value = i; return true;
                    case uint ui: value = ui; return true;
                    case long l: value = l; return true;
                    case ulong ul:
                        if (ul <= long.MaxValue) { value = (long)ul; return true; }
                        break;
                    case float f:
                        value = (long)f; return true;
                    case double d:
                        // Excel의 정수도 double로 들어올 수 있음
                        if (Math.Abs(d % 1.0) < 1e-9) { value = (long)d; return true; }
                        break;
                    case decimal m:
                        if (m == decimal.Truncate(m)) { value = (long)m; return true; }
                        break;
                    case string s:
                        if (long.TryParse(s.Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out var li))
                        { value = li; return true; }
                        break;
                }
            }
            catch { /* ignore */ }

            value = 0;
            return false;
        }

        internal static bool TryToDouble(object v, out double value)
        {
            try
            {
                switch (v)
                {
                    case sbyte sb: value = sb; return true;
                    case byte b: value = b; return true;
                    case short s: value = s; return true;
                    case ushort us: value = us; return true;
                    case int i: value = i; return true;
                    case uint ui: value = ui; return true;
                    case long l: value = l; return true;
                    case ulong ul: value = ul; return true;
                    case float f: value = f; return true;
                    case double d: value = d; return true;
                    case decimal m: value = (double)m; return true;
                    case string s:
                        if (double.TryParse(s.Trim(), NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out var dd))
                        { value = dd; return true; }
                        break;
                }
            }
            catch { /* ignore */ }

            value = 0;
            return false;
        }

        internal static bool TryToBool(object v, out bool value)
        {
            try
            {
                switch (v)
                {
                    case bool b: value = b; return true;
                    case sbyte sb: value = sb != 0; return true;
                    case byte by: value = by != 0; return true;
                    case short s: value = s != 0; return true;
                    case ushort us: value = us != 0; return true;
                    case int i: value = i != 0; return true;
                    case uint ui: value = ui != 0; return true;
                    case long l: value = l != 0; return true;
                    case ulong ul: value = ul != 0; return true;
                    case string str:
                        var t = str.Trim().ToLowerInvariant();
                        if (t is "1" or "true" or "t" or "y" or "yes") { value = true; return true; }
                        if (t is "0" or "false" or "f" or "n" or "no") { value = false; return true; }
                        break;
                }
            }
            catch { /* ignore */ }

            value = false;
            return false;
        }

        internal static bool TryToDate(object v, out DateTime value)
        {
            try
            {
                switch (v)
                {
                    case DateTime dt: value = dt; return true;
                    case double oa:   // Excel OADate
                        value = DateTime.FromOADate(oa);
                        return true;
                    case string s:
                        if (DateTime.TryParse(s.Trim(), CultureInfo.InvariantCulture, DateTimeStyles.AssumeLocal, out var d))
                        { value = d; return true; }
                        break;
                }
            }
            catch { /* ignore */ }

            value = default;
            return false;
        }

        /// <summary>
        /// 문자열로 변환하며 NFC 정규화.
        /// </summary>
        internal static string NormalizeToString(object v)
        {
            var s = v switch
            {
                string ss => ss,
                _ => Convert.ToString(v, CultureInfo.InvariantCulture) ?? string.Empty,
            };
            return s.Normalize(NormalizationForm.FormC);
        }

        // ─────────────────────────────────────────────────────────────────
        //  Log: 워크북에 "_XQL_Log" 시트를 만들어 기록 (UI 스레드에서만 작업)
        //  컬럼: Timestamp | Level | Sheet | Address | Message
        // ─────────────────────────────────────────────────────────────────
        public static void LogInfo(string msg, string? sheet = null, string? address = null) => Log("INFO", msg, sheet, address);
        public static void LogWarn(string msg, string? sheet = null, string? address = null) => Log("WARN", msg, sheet, address);
        public static void LogError(string msg, string? sheet = null, string? address = null) => Log("ERROR", msg, sheet, address);

        public static void Log(string level, string msg, string? sheet, string? address)
        {
            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                Excel.Application app = (Excel.Application)ExcelDnaUtil.Application;
                Excel.Workbook? wb = null; Excel.Worksheet? ws = null;
                Excel.Range? cell = null; Excel.Range? row = null;
                try
                {
                    wb = app.ActiveWorkbook;
                    if (wb == null) return;

                    // 시트 찾기/생성
                    ws = FindOrCreateLogSheet(wb, "_XQL_Log");

                    // 헤더
#pragma warning disable CS8602 // null 가능 참조에 대한 역참조입니다.
                    // UsedRange는 호출 시점 스냅샷. 헤더 작성 전 기준으로 next를 계산해도
                    // 최소 2행부터 쓰도록 Max 처리되어 안전함.
                    var ur = ws.UsedRange as Excel.Range;
                    if ((ur?.Rows?.Count ?? 0) == 1 && ((ws.Cells[1, 1] as Excel.Range)?.Value2 == null))
                    {
                        (ws.Cells[1, 1] as Excel.Range).Value2 = "Timestamp";
                        (ws.Cells[1, 2] as Excel.Range).Value2 = "Level";
                        (ws.Cells[1, 3] as Excel.Range).Value2 = "Sheet";
                        (ws.Cells[1, 4] as Excel.Range).Value2 = "Address";
                        (ws.Cells[1, 5] as Excel.Range).Value2 = "Message";
                    }

                    int last = (ur?.Row ?? 1) + ((ur?.Rows?.Count ?? 1) - 1);
                    int next = Math.Max(2, last + 1);

                    row = ws.Range[ws.Cells[next, 1], ws.Cells[next, 5]];
                    (row.Cells[1, 1] as Excel.Range).Value2 = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
                    (row.Cells[1, 2] as Excel.Range).Value2 = level;
                    (row.Cells[1, 3] as Excel.Range).Value2 = sheet ?? "";
                    (row.Cells[1, 4] as Excel.Range).Value2 = address ?? "";
                    (row.Cells[1, 5] as Excel.Range).Value2 = msg ?? "";
#pragma warning restore CS8602 // null 가능 참조에 대한 역참조입니다.
                    ReleaseCom(ur);
                }
                catch { /* logging failure is non-fatal */ }
                finally
                {
                    ReleaseCom(row); ReleaseCom(cell); ReleaseCom(ws); ReleaseCom(wb);
                }
            });
        }

        private static Excel.Worksheet FindOrCreateLogSheet(Excel.Workbook wb, string name)
        {
            foreach (Excel.Worksheet s in wb.Worksheets)
            {
                try { if (string.Equals(s.Name, name, StringComparison.Ordinal)) return s; }
                finally { ReleaseCom(s); }
            }
            var created = (Excel.Worksheet)wb.Worksheets.Add();
            created.Name = name;
            // 맨 뒤로 이동
            created.Move(After: wb.Worksheets[wb.Worksheets.Count]);
            return created;
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
                    ReleaseCom(ur); ur = null;
                    if (needHeader)
                    {
                        string[] headers = { "Timestamp", "Table", "RowKey", "Column", "Local", "Server", "Type", "Message", "Sheet", "Address" };
                        for (int i = 0; i < headers.Length; i++)
                            (ws.Cells[1, i + 1] as Excel.Range)!.Value2 = headers[i];
                        // 간단 오토필터
                        Excel.Range hdr = ws.Range[ws.Cells[1, 1], ws.Cells[1, headers.Length]];
                        try { ws.ListObjects.Add(Excel.XlListObjectSourceType.xlSrcRange, hdr, Type.Missing, Excel.XlYesNoGuess.xlYes); } catch { }
                        ReleaseCom(hdr);
                    }

                    // 현재 마지막 행
                    ur = ws.UsedRange as Excel.Range;
                    int last = (ur?.Row ?? 1) + ((ur?.Rows?.Count ?? 1) - 1);
                    ReleaseCom(ur); ur = null;

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
                        ReleaseCom(row); row = null;
                    }
                }
                catch { }
                finally
                {
                    ReleaseCom(row); ReleaseCom(ur); ReleaseCom(ws); ReleaseCom(wb);
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

        private static Excel.Worksheet FindOrCreateSheet(Excel.Workbook wb, string name)
        {
            foreach (Excel.Worksheet s in wb.Worksheets)
            {
                try { if (string.Equals(s.Name, name, StringComparison.Ordinal)) return s; }
                finally { ReleaseCom(s); }
            }
            var created = (Excel.Worksheet)wb.Worksheets.Add();
            created.Name = name;
            created.Move(After: wb.Worksheets[wb.Worksheets.Count]);
            return created;
        }
    }
}