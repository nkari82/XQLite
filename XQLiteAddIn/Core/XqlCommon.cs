// XqlCommon.cs
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;


namespace XQLite.AddIn
{
    internal static class XqlCommon
    {
        // XqlCommon.cs (공용으로 쓰고 싶다면)
        public static class Monotonic
        {
            [MethodImpl(MethodImplOptions.AggressiveInlining)]
            public static long NowMs() => (Stopwatch.GetTimestamp() * 1000L) / Stopwatch.Frequency;
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

        // ── 파일/상태 보조 ──────────────────────────────────────────────
        /// <summary>워크북 경로 옆에 숨김 상태폴더 생성/보장 (기본 ".xql")</summary>
        public static string EnsureHiddenStateDir(string workbookFullName, string? dirName = null)
        {
            string baseDir = Path.GetDirectoryName(workbookFullName)
                             ?? Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            var leaf = string.IsNullOrWhiteSpace(dirName) ? XqlConfig.StateDirName : dirName!;
            if (string.IsNullOrWhiteSpace(leaf)) leaf = ".xql";
            string dir = Path.Combine(baseDir, leaf);
            try
            {
                if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);
                var di = new DirectoryInfo(dir);
                di.Attributes |= FileAttributes.Hidden; // 윈도우 숨김
            }
            catch { /* 무시 */ }
            return dir;
        }

        public static T? LoadJsonFile<T>(string path)
        {
            try { if (!File.Exists(path)) return default; return JsonConvert.DeserializeObject<T>(File.ReadAllText(path)); }
            catch { return default; }
        }

        public static void SaveJsonFile<T>(string path, T data)
        {
            try
            {
                var json = JsonConvert.SerializeObject(data, Formatting.Indented);
                File.WriteAllText(path, json);
            }
            catch { /* 조용히 무시 */ }
        }

        public static string SanitizeFileStem(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return "default";
            var bad = Path.GetInvalidFileNameChars();
            var arr = s.Select(ch => bad.Contains(ch) ? '_' : ch).ToArray();
            return new string(arr);
        }
    }
}