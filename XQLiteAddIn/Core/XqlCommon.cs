// XqlCommon.cs
using ExcelDna.Integration;
using System;
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

        private static int _excelMainThreadId = -1;

        static XqlCommon()
        {
            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                try { _excelMainThreadId = Thread.CurrentThread.ManagedThreadId; }
                catch { }
            });
        }


        /// <summary>
        /// UI → BG → UI 세 단계 파이프라인을 안전하게 수행합니다.
        /// - captureOnUi : Excel UI 스레드에서 순수 스냅샷(TSnap)만 캡처 (COM RCW 보관 금지)
        /// - workOnBg    : 백그라운드에서 스냅샷을 이용해 순수 연산 수행(네트워크/CPU 등)
        /// - applyOnUi   : 결과를 UI에 반영 (필요 시 Excel 개체 접근)
        /// CancellationToken은 각 hop 사이에서 체크되어 즉시 중단됩니다.
        /// </summary>
        public static async Task BridgeAsync<TSnap, TResult>(
            Func<TSnap> captureOnUi,
            Func<TSnap, CancellationToken, Task<TResult>> workOnBg,
            Action<TResult> applyOnUi,
            CancellationToken ct = default)
        {
            // --- 인자 검증 ---
            if (captureOnUi is null) throw new ArgumentNullException(nameof(captureOnUi));
            if (workOnBg is null) throw new ArgumentNullException(nameof(workOnBg));
            if (applyOnUi is null) throw new ArgumentNullException(nameof(applyOnUi));

            // 빠른 취소
            ct.ThrowIfCancellationRequested();

            TSnap snap;
            try
            {
                // UI hop (Excel UI 스레드로 마샬링)
                snap = await OnExcelThreadAsync(captureOnUi).ConfigureAwait(false);
            }
            catch (OperationCanceledException) { throw; }
            catch (Exception ex)
            {
                // 캡처 단계 예외는 그대로 전파하되 로깅
                try { XqlLog.Warn("BridgeAsync.captureOnUi failed: " + ex.Message); } catch { }
                throw;
            }

            // hop 사이에서도 취소 체크
            ct.ThrowIfCancellationRequested();

            TResult result;
            try
            {
                // BG hop (절대 COM 접근 금지; 네트워크/CPU 위주)
                result = await workOnBg(snap, ct).ConfigureAwait(false);
            }
            catch (OperationCanceledException) { throw; }
            catch (Exception ex)
            {
                // 백그라운드 작업 예외는 로깅 후 전파
                try { XqlLog.Warn("BridgeAsync.workOnBg failed: " + ex.Message); } catch { }
                throw;
            }

            // apply 직전 취소면 UI 반영 스킵
            ct.ThrowIfCancellationRequested();

            try
            {
                // UI hop (결과 반영)
                await OnExcelThreadAsync(() =>
                {
                    applyOnUi(result);
                    return 0;
                }).ConfigureAwait(false);
            }
            catch (OperationCanceledException) { throw; }
            catch (Exception ex)
            {
                // 반영 단계 예외 로깅 후 전파
                try { XqlLog.Warn("BridgeAsync.applyOnUi failed: " + ex.Message); } catch { }
                throw;
            }
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

        // XqlCommon.cs
        internal static void ReleaseCom(params object?[] objs)
        {
            foreach (var o in objs)
            {
                try
                {
                    if (o == null) continue;
                    if (!Marshal.IsComObject(o)) continue;

                    // 강제 분리(FinalRelease) 금지 — 참조 1회만 안전하게 감소
                    try
                    {
                        Marshal.ReleaseComObject(o);
                    }
                    catch (InvalidComObjectException)
                    {
                        // 이미 분리된 RCW — 무시
                    }
                }
                catch (COMException)
                {
                    // Excel 종료/분리 중 예외 — 무시
                }
                catch
                {
                    // 어떤 예외도 삼킴
                }
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

        // 숫자/일반용 Clamp (generic)
        public static T Clamp<T>(T value, T min, T max) where T : IComparable<T>
        {
            if (min.CompareTo(max) > 0) (min, max) = (max, min); // 잘못 준 경우 보정
            if (value.CompareTo(min) < 0) return min;
            if (value.CompareTo(max) > 0) return max;
            return value;
        }

        // Excel 셀의 값/서식 조합으로 DateTime 가능성 판정
        internal static bool IsExcelDateTimeLikely(Excel.Range c)
        {
            try
            {
                if (c == null) return false;
                var v = c.Value2;

                // 값이 숫자(OA Date로 해석 가능)인지 확인
                if (v is double d)
                {
                    // 흔한 날짜 범위 (1900-01-01 ~ 9999-12-31 근사)
                    // Excel OA: 1 = 1899-12-31, 2 = 1900-01-01, ... (윤년 버그 고려 여유 범위)
                    if (d >= -657434 && d <= 2958465) // DateTime.OADate 허용범위
                        return true;
                }

                // 날짜/시간 서식 여부 체크(영문/로컬)
                string? nf = null, nfl = null;
                try { nf = Convert.ToString(c.NumberFormat); } catch { }
                try { nfl = Convert.ToString(c.NumberFormatLocal); } catch { }
                string s = ((nf ?? "") + ";" + (nfl ?? "")).ToLowerInvariant();

                // 전형적인 날짜/시간 토큰 포함 여부 (m/d/y/h/s 등)
                if (s.Contains("yy") || s.Contains("yyyy") || s.Contains("m/") || s.Contains("mm") ||
                    s.Contains("dd") || s.Contains("d ") || s.Contains("h:") || s.Contains("hh") ||
                    s.Contains("ampm") || s.Contains("오전") || s.Contains("오후"))
                    return true;

                return false;
            }
            catch { return false; }
        }


        public readonly struct ExcelBatchScope : IDisposable
        {
            private readonly Excel.Application? _app;

            // 캡처에 성공한 항목만 복구
            private readonly bool _capEvents, _capScreen, _capAlerts;
            private readonly bool _oldEvents, _oldScreen, _oldAlerts;

            // Calculation은 값을 "읽지 않는다"
            private readonly bool _calcTouched; // 수동 전환 성공 여부만 기억

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

                // 1) 상태 캡처 (읽기 안전한 것만)
                try { _oldEvents = app.EnableEvents; _capEvents = true; } catch { }
                try { _oldScreen = app.ScreenUpdating; _capScreen = true; } catch { }
                try { _oldAlerts = app.DisplayAlerts; _capAlerts = true; } catch { }

                // 2) 배치 모드 진입 (가능한 것만)
                try { app.EnableEvents = false; } catch { }
                try { app.ScreenUpdating = false; } catch { }
                try { app.DisplayAlerts = false; } catch { }

                // Calculation은 "읽지 않고" 바로 수동으로 시도
                try
                {
                    app.Calculation = Excel.XlCalculation.xlCalculationManual;
                    _calcTouched = true; // 전환 성공
                }
                catch
                {
                    // 전환 실패: 만지지 않음(복구도 시도하지 않음)
                    _calcTouched = false;
                }
            }

            public void Dispose()
            {
                if (_app is null) return;

                // 1) Calculation 복귀: 값을 읽어두지 않았으므로 '자동'으로만 되돌린다.
                //    (수동 전환에 성공했을 때만 복귀 시도)
                if (_calcTouched)
                {
                    try { _app.Calculation = Excel.XlCalculation.xlCalculationAutomatic; } catch { }
                }

                // 2) 나머지는 캡처 성공한 항목만 복구
                try { if (_capAlerts) _app.DisplayAlerts = _oldAlerts; } catch { }
                try { if (_capScreen) _app.ScreenUpdating = _oldScreen; } catch { }
                try { if (_capEvents) _app.EnableEvents = _oldEvents; } catch { }
            }
        }


        /// <summary>
        /// COM 객체를 스코프 기반으로 안전하게 관리하기 위한 래퍼.
        /// - Dispose 시 ReleaseComObject 1회만 호출 (FinalRelease 금지)
        /// - Detach()로 소유권 이전 가능
        /// - Use/Map 헬퍼 제공
        /// - await 경계 넘기지 말 것 (스코프 생명주기 보장)
        /// - 스레드 제약 없음 (사용자 책임)
        /// </summary>
        public sealed class SmartCom<T> : IDisposable where T : class
        {
            private T? _obj;
            private bool _disposed;

            public SmartCom(T? obj)
            {
                _obj = obj;
            }

            /// <summary>래핑된 실제 객체(읽기전용).</summary>
            public T? Value => _disposed ? null : _obj;

            /// <summary>암시적 변환: SmartCom<T>; → T?</summary>
            public static implicit operator T?(SmartCom<T> w) => w?._obj;

            /// <summary>소유권을 호출자에게 넘기고 이 래퍼는 해제하지 않음.</summary>
            public T? Detach()
            {
                var o = Interlocked.Exchange(ref _obj, null);
                return o;
            }

            /// <summary>사용 후 자동 해제 (Action 버전)</summary>
            public void Use(Action<T> action)
            {
                if (_disposed) throw new ObjectDisposedException(nameof(SmartCom<T>));
                var o = _obj;
                if (o == null) { Dispose(); return; }
                try { action(o); }
                finally { Dispose(); }
            }

            /// <summary>사용 후 자동 해제 + 결과 반환 (Func 버전)</summary>
            public R? Map<R>(Func<T, R?> func)
            {
                if (_disposed) throw new ObjectDisposedException(nameof(SmartCom<T>));
                var o = _obj;
                if (o == null) { Dispose(); return default; }
                try { return func(o); }
                finally { Dispose(); }
            }

            /// <summary>해제: ReleaseComObject 1회만 호출 (RCW 분리 금지)</summary>
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

            // -------- Factory ----------
            public static SmartCom<T> Wrap(T? obj) => new SmartCom<T>(obj);

            public static SmartCom<T> Wrap(object? obj) => new SmartCom<T>(obj as T);

            public static SmartCom<T> Acquire(Func<T?> acquire) => new SmartCom<T>(acquire());
        }

        /// <summary>
        /// 여러 COM 객체를 한 스코프에서 관리 (역순 해제)
        /// - Add()/Get() 으로 등록
        /// - Dispose 시 ReleaseComObject 1회씩
        /// - 스레드 제약 없음
        /// </summary>
        public sealed class SmartComBatch : IDisposable
        {
            private readonly List<object?> _objs = new(32);
            private bool _disposed;

            /// <summary>등록 후 SmartCom 형태로 반환</summary>
            public SmartCom<T> Get<T>(Func<T?> acquire) where T : class
            {
                var o = acquire();
                if (o != null) _objs.Add(o);
                return SmartCom<T>.Wrap(o);
            }

            /// <summary>이미 얻은 COM 객체를 등록</summary>
            public void Add(object? o)
            {
                if (o != null) _objs.Add(o);
            }

            /// <summary>해당 객체를 해제 목록에서 제외(소유권 이전)</summary>
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
                        if (o != null && System.Runtime.InteropServices.Marshal.IsComObject(o))
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(o);
                    }
                    catch (System.Runtime.InteropServices.InvalidComObjectException) { }
                    catch (System.Runtime.InteropServices.COMException) { }
                    catch { }
                }
                _objs.Clear();
            }
        }


        // ───────────────────────── Mso Image Helper ─────────────────────────
        public static class MsoImageHelper
        {
            /// <summary>Office 버전/PIA 차이를 피하기 위해 리플렉션으로 GetImageMso 호출.</summary>
            public static stdole.IPictureDisp? Get(string idMso, int size = 32)
            {
                if (string.IsNullOrWhiteSpace(idMso)) return null;

                stdole.IPictureDisp? pic;
                try
                {
                    object app = ExcelDnaUtil.Application;                // Excel.Application (COM)
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

                    pic = ret as stdole.IPictureDisp;

                    return pic;
                }
                catch
                {
                    // 실패 시 null 반환 → UI는 기본 아이콘/텍스트로 진행
                    return null;
                }
            }
        }

    }
}