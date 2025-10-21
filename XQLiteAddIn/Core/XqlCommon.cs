// XqlCommon.cs
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;


namespace XQLite.AddIn
{
    internal static class XqlCommon
    {

        private static int _excelMainThreadId = -1;

        static XqlCommon()
        {
            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                try { _excelMainThreadId = Thread.CurrentThread.ManagedThreadId; }
                catch {}
            });
        }

        /// <summary>
        /// Excel UI 쓰레드에서 동기 작업을 실행하고 결과를 Task로 돌려준다.
        /// (work 안에서는 반드시 COM 개체를 ReleaseCom로 정리할 것)
        /// </summary>

        public static bool IsExcelMainThread()
        {
            return Thread.CurrentThread.ManagedThreadId == _excelMainThreadId;
        }

        public static Task<T> OnExcelThreadAsync<T>(
            Func<T> work,
            int? timeoutMs = null,
            CancellationToken ct = default)
        {
            if (work == null) throw new ArgumentNullException(nameof(work));

            // UI thread면 즉시 수행(가장 빠른 경로)
            if (IsExcelMainThread())
            {
                try { return Task.FromResult(work()); }
                catch (Exception ex) { return Task.FromException<T>(ex); }
            }

            var tcs = new TaskCompletionSource<T>(TaskCreationOptions.RunContinuationsAsynchronously);

            // timeout / cancel
            CancellationTokenSource? cts = null;
            var token = ct;
            if (timeoutMs is int ms && ms > 0)
            {
                cts = CancellationTokenSource.CreateLinkedTokenSource(ct);
                cts.CancelAfter(ms);
                token = cts.Token;
            }
            var reg = token.CanBeCanceled ? token.Register(() => tcs.TrySetCanceled(token)) : default;

            // 완료 후 정리
            tcs.Task.ContinueWith(_ => { try { reg.Dispose(); } catch { } try { cts?.Dispose(); } catch { } },
                                  TaskScheduler.Default);

            try
            {
                ExcelAsyncUtil.QueueAsMacro(() =>
                {
                    if (token.IsCancellationRequested || tcs.Task.IsCompleted) return;

                    try
                    {
                        var r = work();
                        tcs.TrySetResult(r);
                    }
                    catch (Exception ex)
                    {
                        tcs.TrySetException(ex);
                    }
                });
            }
            catch (Exception qex)
            {
                tcs.TrySetException(qex);
            }

            return tcs.Task;
        }

        public static Task OnExcelThreadAsync(Action work, int? timeoutMs = null, CancellationToken ct = default)
            => OnExcelThreadAsync<object?>(() => { work(); return null; }, timeoutMs, ct);
        
        public static long NowMs() => (Stopwatch.GetTimestamp() * 1000L) / Stopwatch.Frequency;


        /// <summary>값 정규화(전송/비교용) – 모든 모듈에서 이 함수만 사용.</summary>
        public static string? Canonicalize(object? v)
        {
            if (v is null) return null;
            switch (v)
            {
                case bool b: return b ? "1" : "0";
                case double d: return d.ToString("R", CultureInfo.InvariantCulture);
                case float f: return ((double)f).ToString("R", CultureInfo.InvariantCulture);
                case int i: return i.ToString(CultureInfo.InvariantCulture);
                case long l: return l.ToString(CultureInfo.InvariantCulture);
                case decimal m: return ((double)m).ToString("R", CultureInfo.InvariantCulture);
                case DateTime dt:
                    var ms = (long)(dt.ToUniversalTime() - new DateTime(1970, 1, 1)).TotalMilliseconds;
                    return ms.ToString(CultureInfo.InvariantCulture);
                default:
                    var s = v.ToString();
                    return string.IsNullOrWhiteSpace(s) ? "" : s!;
            }
        }

        // SHA-1 64비트 축약 지문(16 hex)
        internal static string Fingerprint(object? v)
        {
            var s = Canonicalize(v);
            using var sha1 = SHA1.Create();
            var bytes = sha1.ComputeHash(Encoding.UTF8.GetBytes(s));
            ulong u = ((ulong)bytes[0] << 56) | ((ulong)bytes[1] << 48) | ((ulong)bytes[2] << 40) | ((ulong)bytes[3] << 32)
                    | ((ulong)bytes[4] << 24) | ((ulong)bytes[5] << 16) | ((ulong)bytes[6] << 8) | bytes[7];
            return u.ToString("x16");
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

        public static void ReleaseCom(params object?[] objs)
        {
            foreach (var o in objs)
            {
                try
                {
                    if (o == null) continue;
                    if (!Marshal.IsComObject(o)) continue;
                    // 여러 번 Release 호출돼도 0 미만으로 내려가는 일은 없습니다.
                    Marshal.FinalReleaseComObject(o);
                }
                catch (COMException) { /* Excel 종료/분리 중 */ }
                catch (InvalidComObjectException) { /* RCW 분리됨 */ }
                catch { /* no-op */ }
            }
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
    }
}