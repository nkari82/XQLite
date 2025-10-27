// XqlCommon.cs (cleaned & production-hardened)
using ExcelDna.Integration;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Reflection;
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
        static XqlCommon()
        {
        }

        /// <summary>
        /// Excel UI 스레드에서 동기 delegate를 실행하여 결과를 Task로 돌려줍니다.
        /// - UI 스레드인 경우 즉시 실행(빠른 경로)
        /// - 비-UI 스레드 → QueueAsMacro로 hop
        /// - work 내부에서 획득한 COM은 반드시 ReleaseCom로 정리하세요.
        /// </summary>
        public static Task<T> OnExcelThreadAsync<T>(
            Func<T> work,
            int? timeoutMs = null,
            CancellationToken ct = default)
        {
            if (work == null) throw new ArgumentNullException(nameof(work));

            var tcs = new TaskCompletionSource<T>(TaskCreationOptions.RunContinuationsAsynchronously);

            // timeout / cancel 설정
            CancellationTokenSource? linkedCts = null;
            var token = ct;
            if (timeoutMs is int ms && ms > 0)
            {
                linkedCts = CancellationTokenSource.CreateLinkedTokenSource(ct);
                linkedCts.CancelAfter(ms);
                token = linkedCts.Token;
            }

            var reg = token.CanBeCanceled ? token.Register(() => tcs.TrySetCanceled(token)) : default;

            // 완료/실패 후 등록 해제/CTS 정리
            tcs.Task.ContinueWith(_ =>
            {
                try { reg.Dispose(); } catch { }
                try { linkedCts?.Dispose(); } catch { }
            }, TaskScheduler.Default);

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

        /// <summary>
        /// UI → BG → UI 세 단계 파이프라인을 안전하게 수행합니다.
        /// - captureOnUi : Excel UI 스레드에서 스냅샷(TSnap) 캡처 (COM RCW 보관 금지)
        /// - workOnBg    : 백그라운드(절대 COM 접근 금지)에서 스냅샷 기반 연산
        /// - applyOnUi   : 결과를 UI에 반영 (Excel 개체 접근 가능)
        /// 각 단계 사이에 CancellationToken을 체크합니다.
        /// </summary>
        public static async Task BridgeAsync<TSnap, TResult>(
            Func<TSnap> captureOnUi,
            Func<TSnap, CancellationToken, Task<TResult>> workOnBg,
            Action<TResult> applyOnUi,
            CancellationToken ct = default)
        {
            if (captureOnUi is null) throw new ArgumentNullException(nameof(captureOnUi));
            if (workOnBg is null) throw new ArgumentNullException(nameof(workOnBg));
            if (applyOnUi is null) throw new ArgumentNullException(nameof(applyOnUi));

            ct.ThrowIfCancellationRequested();

            TSnap snap;
            try
            {
                snap = await OnExcelThreadAsync(captureOnUi).ConfigureAwait(false);
            }
            catch (OperationCanceledException) { throw; }
            catch (Exception ex)
            {
                try { XqlLog.Warn("BridgeAsync.captureOnUi failed: " + ex.Message); } catch { }
                throw;
            }

            ct.ThrowIfCancellationRequested();

            TResult result;
            try
            {
                result = await workOnBg(snap, ct).ConfigureAwait(false);
            }
            catch (OperationCanceledException) { throw; }
            catch (Exception ex)
            {
                try { XqlLog.Warn("BridgeAsync.workOnBg failed: " + ex.Message); } catch { }
                throw;
            }

            ct.ThrowIfCancellationRequested();

            try
            {
                await OnExcelThreadAsync(() =>
                {
                    applyOnUi(result);
                    return 0;
                }).ConfigureAwait(false);
            }
            catch (OperationCanceledException) { throw; }
            catch (Exception ex)
            {
                try { XqlLog.Warn("BridgeAsync.applyOnUi failed: " + ex.Message); } catch { }
                throw;
            }
        }

        // =====================================================================
        // Time / Math
        // =====================================================================

        public static long NowMs() => (Stopwatch.GetTimestamp() * 1000L) / Stopwatch.Frequency;

        public static T Clamp<T>(T value, T min, T max) where T : IComparable<T>
        {
            if (min.CompareTo(max) > 0) (min, max) = (max, min);
            if (value.CompareTo(min) < 0) return min;
            if (value.CompareTo(max) > 0) return max;
            return value;
        }

        public static void InterlockedMax(ref long target, long value)
        {
            while (true)
            {
                long cur = Volatile.Read(ref target);
                if (value <= cur) return;
                if (Interlocked.CompareExchange(ref target, value, cur) == cur) return;
            }
        }

        // =====================================================================
        // String / CSV / Hash
        // =====================================================================

        /// <summary>SHA-1 64비트 축약 지문(16 hex)</summary>
        internal static string Fingerprint(object? v)
        {
            var s = Canonicalize(v) ?? string.Empty;
            using var sha1 = SHA1.Create();
            var bytes = sha1.ComputeHash(Encoding.UTF8.GetBytes(s));
            ulong u = ((ulong)bytes[0] << 56) | ((ulong)bytes[1] << 48) | ((ulong)bytes[2] << 40) | ((ulong)bytes[3] << 32)
                    | ((ulong)bytes[4] << 24) | ((ulong)bytes[5] << 16) | ((ulong)bytes[6] << 8) | bytes[7];
            return u.ToString("x16");
        }

        /// <summary>Excel Column Index -> "A..Z, AA.."</summary>
        internal static string ColumnIndexToLetter(int col)
        {
            if (col <= 0) return "A";
            string s = string.Empty;
            while (col > 0)
            {
                int m = (col - 1) % 26;
                s = (char)('A' + m) + s;
                col = (col - 1) / 26;
            }
            return s;
        }

        internal static string CsvEscape(string? s)
        {
            s ??= "";
            bool needQuote = s.Contains(',') || s.Contains('"') || s.Contains('\n') || s.Contains('\r');
            return needQuote ? "\"" + s.Replace("\"", "\"\"") + "\"" : s;
        }

        public static string? ValueToString(object? v)
        {
            if (v == null || v is DBNull) return null;

            switch (v)
            {
                case string s:
                    return s; // (원본 그대로; 필요시 여기서 Normalize(FormC) 추가 가능)

                case bool b:
                    return b ? "true" : "false";

                case DateTimeOffset dto:
                    // ISO 8601 with offset (서버 텍스트 컬럼과 일관)
                    return dto.ToString("yyyy-MM-ddTHH:mm:ss.fffK", CultureInfo.InvariantCulture);

                case DateTime dt:
                    // Kind 보존해 ISO 문자열
                    return dt.ToString("yyyy-MM-ddTHH:mm:ss.fffK", CultureInfo.InvariantCulture);

                case double or float or decimal or long or int or short or byte or sbyte or uint or ulong:
                    return Convert.ToString(v, CultureInfo.InvariantCulture);

                default:
                    return Convert.ToString(v, CultureInfo.InvariantCulture);
            }
        }

        public static string? Canonicalize(object? v)
        {
            if (v == null || v is DBNull) return null;

            if (v is DateTimeOffset dto)
                return dto.ToString("yyyy-MM-ddTHH:mm:ss.fffK", CultureInfo.InvariantCulture);

            if (v is DateTime dt)
                return dt.ToString("yyyy-MM-ddTHH:mm:ss.fffK", CultureInfo.InvariantCulture);

            return ValueToString(v);
        }

        internal static bool EqualKey(object? a, object? b)
        {
            if (a is null && b is null) return true;
            if (a is null || b is null) return false;

            var sa = (Convert.ToString(a, CultureInfo.InvariantCulture) ?? "").Trim().Normalize(NormalizationForm.FormC);
            var sb = (Convert.ToString(b, CultureInfo.InvariantCulture) ?? "").Trim().Normalize(NormalizationForm.FormC);
            return string.Equals(sa, sb, StringComparison.Ordinal);
        }


        internal static bool IsNullish(object? v)
        {
            if (v is null || v is DBNull) return true;
            if (v is string s) return string.IsNullOrWhiteSpace(s);
            return false;
        }

        // =====================================================================
        // Conversions
        // =====================================================================

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
                    case ulong ul: if (ul <= long.MaxValue) { value = (long)ul; return true; } break;
                    case float f: value = (long)f; return true;
                    case double d: if (Math.Abs(d % 1.0) < 1e-9) { value = (long)d; return true; } break;
                    case decimal m: if (m == decimal.Truncate(m)) { value = (long)m; return true; } break;
                    case string str:
                        if (long.TryParse(str.Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out var li))
                        { value = li; return true; }
                        break;
                }
            }
            catch { }
            value = 0; return false;
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
                    case string str:
                        if (double.TryParse(str.Trim(), NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out var dd))
                        { value = dd; return true; }
                        break;
                }
            }
            catch { }
            value = 0; return false;
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
            catch { }
            value = false; return false;
        }

        internal static bool TryToDate(object v, out DateTime value)
        {
            try
            {
                switch (v)
                {
                    case DateTime dt:
                        value = dt; return true;

                    case DateTimeOffset dto:
                        value = dto.LocalDateTime; return true;

                    case double oa:
                        value = DateTime.FromOADate(oa); return true;

                    case string s:
                        var t = s.Trim();

                        // ISO 8601을 먼저 시도 (오탐 최소화)
                        if (DateTimeOffset.TryParse(t, CultureInfo.InvariantCulture,
                                DateTimeStyles.AssumeLocal | DateTimeStyles.AllowWhiteSpaces, out var dto2))
                        { value = dto2.LocalDateTime; return true; }

                        if (DateTime.TryParse(t, CultureInfo.InvariantCulture,
                                DateTimeStyles.AssumeLocal | DateTimeStyles.AllowWhiteSpaces, out var d))
                        { value = d; return true; }
                        break;
                }
            }
            catch { }
            value = default; return false;
        }


        /// <summary>문자열로 변환하며 NFC 정규화.</summary>
        internal static string NormalizeToString(object v)
        {
            var s = v is string ss ? ss : (Convert.ToString(v, CultureInfo.InvariantCulture) ?? string.Empty);
            return s.Normalize(NormalizationForm.FormC);
        }

        // =====================================================================
        // Excel helpers
        // =====================================================================

        /// <summary>Excel 셀의 값/서식 조합으로 DateTime 가능성 판정</summary>
        internal static bool IsExcelDateTimeLikely(Excel.Range c)
        {
            try
            {
                if (c == null) return false;
                var v = c.Value2;

                if (v is double d)
                {
                    // DateTime.OADate 허용범위
                    if (d >= -657434 && d <= 2958465) return true;
                }

                string? nf = null, nfl = null;
                try { nf = Convert.ToString(c.NumberFormat, CultureInfo.InvariantCulture); } catch { }
                try { nfl = Convert.ToString(c.NumberFormatLocal, CultureInfo.InvariantCulture); } catch { }
                string s = ((nf ?? "") + ";" + (nfl ?? "")).ToLowerInvariant();

                if (s.Contains("yy") || s.Contains("yyyy") || s.Contains("m/") || s.Contains("mm")
                    || s.Contains("dd") || s.Contains("d ") || s.Contains("h:") || s.Contains("hh")
                    || s.Contains("ampm") || s.Contains("오전") || s.Contains("오후"))
                    return true;

                return false;
            }
            catch { return false; }
        }

        internal static Excel.Range? IntersectSafe(Excel.Worksheet ws, Excel.Range a, Excel.Range b)
        {
            try { return ws.Application.Intersect(a, b); }
            catch { return null; }
        }

        /// <summary>
        /// Excel 화면/이벤트/알림을 일시적으로 끄고, 계산을 수동으로 전환하는 배치 스코프.
        /// 캡처에 성공한 속성만 복구합니다.
        /// </summary>
        public readonly struct ExcelBatchScope : IDisposable
        {
            private readonly Excel.Application? _app;
            private readonly bool _capEvents, _capScreen, _capAlerts;
            private readonly bool _oldEvents, _oldScreen, _oldAlerts;
            private readonly bool _calcTouched;

            public ExcelBatchScope(Excel.Application? app)
            {
                _app = app;

                _oldEvents = _oldScreen = _oldAlerts = false;
                _capEvents = _capScreen = _capAlerts = false;
                _calcTouched = false;

                if (app == null || app.Workbooks.Count == 0)
                {
                    _calcTouched = false;
                    return;
                }

                // 1) 상태 캡처
                try { _oldEvents = app.EnableEvents; _capEvents = true; } catch { }
                try { _oldScreen = app.ScreenUpdating; _capScreen = true; } catch { }
                try { _oldAlerts = app.DisplayAlerts; _capAlerts = true; } catch { }

                // 2) 배치 모드
                try { app.EnableEvents = false; } catch { }
                try { app.ScreenUpdating = false; } catch { }
                try { app.DisplayAlerts = false; } catch { }

                try
                {
                    app.Calculation = Excel.XlCalculation.xlCalculationManual;
                    _calcTouched = true;
                }
                catch
                {
                    _calcTouched = false;
                }
            }

            public void Dispose()
            {
                if (_app is null) return;

                if (_calcTouched)
                {
                    try { _app.Calculation = Excel.XlCalculation.xlCalculationAutomatic; } catch { }
                }

                try { if (_capAlerts) _app.DisplayAlerts = _oldAlerts; } catch { }
                try { if (_capScreen) _app.ScreenUpdating = _oldScreen; } catch { }
                try { if (_capEvents) _app.EnableEvents = _oldEvents; } catch { }
            }
        }

        // =====================================================================
        // COM utils
        // =====================================================================

        internal static void ReleaseCom(params object?[] objs)
        {
            foreach (var o in objs)
            {
                try
                {
                    if (o == null) continue;
                    if (!Marshal.IsComObject(o)) continue;

                    try { Marshal.ReleaseComObject(o); }
                    catch (InvalidComObjectException) { }
                }
                catch (COMException) { }
                catch { }
            }
        }

        /// <summary>
        /// COM 객체를 스코프 기반으로 안전하게 관리하기 위한 래퍼.
        /// - Dispose 시 ReleaseComObject 1회만 호출 (FinalRelease 금지)
        /// - Detach()로 소유권 이전 가능
        /// - await 경계 넘기지 말 것
        /// </summary>
        public sealed class SmartCom<T> : IDisposable where T : class
        {
            private T? _obj;
            private bool _disposed;

            public SmartCom(T? obj) { _obj = obj; }

            public T? Value => _disposed ? null : _obj;

            public static implicit operator T?(SmartCom<T> w) => w?._obj;

            public T? Detach()
            {
                var o = Interlocked.Exchange(ref _obj, null);
                return o;
            }

            public void Use(Action<T> action)
            {
                if (_disposed) throw new ObjectDisposedException(nameof(SmartCom<T>));
                var o = _obj;
                if (o == null) { Dispose(); return; }
                try { action(o); }
                finally { Dispose(); }
            }

            public R? Map<R>(Func<T, R?> func)
            {
                if (_disposed) throw new ObjectDisposedException(nameof(SmartCom<T>));
                var o = _obj;
                if (o == null) { Dispose(); return default; }
                try { return func(o); }
                finally { Dispose(); }
            }

            public void Dispose()
            {
                if (_disposed) return;
                _disposed = true;

                var o = Interlocked.Exchange(ref _obj, null);
                if (o == null) return;

                try
                {
                    if (Marshal.IsComObject(o))
                        Marshal.ReleaseComObject(o);
                }
                catch (InvalidComObjectException) { }
                catch (COMException) { }
                catch { }
            }

            public static SmartCom<T> Wrap(T? obj) => new(obj);
            public static SmartCom<T> Wrap(object? obj) => new(obj as T);
            public static SmartCom<T> Acquire(Func<T?> acquire) => new(acquire());
        }

        /// <summary>
        /// 여러 COM 객체를 한 스코프에서 관리 (역순 해제)
        /// </summary>
        public sealed class SmartComBatch : IDisposable
        {
            private readonly List<object?> _objs = new(32);
            private bool _disposed;

            public SmartCom<T> Get<T>(Func<T?> acquire) where T : class
            {
                var o = acquire();
                if (o != null) _objs.Add(o);
                return SmartCom<T>.Wrap(o);
            }

            public void Add(object? o)
            {
                if (o != null) _objs.Add(o);
            }

            public bool Detach(object? o)
            {
                if (o == null) return false;
                return _objs.Remove(o);
            }

            public void Dispose()
            {
                if (_disposed) return;
                _disposed = true;

                for (int i = _objs.Count - 1; i >= 0; --i)
                {
                    var o = _objs[i];
                    try
                    {
                        if (o != null && Marshal.IsComObject(o))
                            Marshal.ReleaseComObject(o);
                    }
                    catch (InvalidComObjectException) { }
                    catch (COMException) { }
                    catch { }
                }
                _objs.Clear();
            }
        }

        // =====================================================================
        // File / Zip
        // =====================================================================

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

        // =====================================================================
        // Collections
        // =====================================================================

        internal static IEnumerable<List<T>> Chunk<T>(IReadOnlyList<T> list, int size)
        {
            if (size <= 0) size = 1000;
            for (int i = 0; i < list.Count; i += size)
                yield return list.Skip(i).Take(Math.Min(size, list.Count - i)).ToList();
        }

        // =====================================================================
        // Mso Image Helper (with simple cache)
        // =====================================================================

        public static class MsoImageHelper
        {
            private static readonly ConcurrentDictionary<(string id, int size), stdole.IPictureDisp?> _cache =
                new();

            /// <summary>Office 버전/PIA 차이를 피하기 위해 리플렉션으로 GetImageMso 호출.</summary>
            public static stdole.IPictureDisp? Get(string idMso, int size = 32)
            {
                if (string.IsNullOrWhiteSpace(idMso)) return null;
                var key = (idMso, size);
                if (_cache.TryGetValue(key, out var cached)) return cached;

                try
                {
                    object app = ExcelDnaUtil.Application; // Excel.Application (COM)
                    var appType = app.GetType();

                    // Application.CommandBars
                    object commandBars = appType.InvokeMember(
                        "CommandBars",
                        BindingFlags.GetProperty,
                        binder: null, target: app, args: null);

                    // CommandBars.GetImageMso(string idMso, int width, int height) : IPictureDisp
                    object? ret = commandBars.GetType().InvokeMember(
                        "GetImageMso",
                        BindingFlags.InvokeMethod,
                        binder: null, target: commandBars,
                        args: new object[] { idMso, size, size });

                    var pic = ret as stdole.IPictureDisp;
                    _cache[key] = pic;
                    return pic;
                }
                catch
                {
                    _cache[key] = null;
                    return null;
                }
            }
        }
    }
}
