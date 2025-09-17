using System;
using System.Threading;
using System.Threading.Tasks;

namespace XQLite.AddIn
{
    public static class XqlNetPolicy
    {
        private static int _failCount;
        private static DateTime _halfOpenAt = DateTime.MinValue;

        // 브레이커: 실패 5회 → 15초 차단, 이후 반개방(1회 시도)
        public static bool CanSend()
        {
            if (_failCount < 5) return true;
            if (DateTime.UtcNow < _halfOpenAt) return false;
            // half-open: allow one attempt
            return true;
        }

        public static void OnSuccess()
        {
            _failCount = 0; _halfOpenAt = DateTime.MinValue;
        }

        public static void OnFailure()
        {
            _failCount++;
            if (_failCount == 5) _halfOpenAt = DateTime.UtcNow.AddSeconds(15);
        }

        public static async Task<T> WithRetryAsync<T>(Func<Task<T>> action, CancellationToken ct = default)
        {
            int attempt = 0;
            Exception? last = null;
            while (attempt < 5)
            {
                ct.ThrowIfCancellationRequested();
                if (!CanSend()) { await Task.Delay(1000, ct); continue; }
                try
                {
                    var result = await action();
                    OnSuccess();
                    return result;
                }
                catch (Exception ex)
                {
                    last = ex; OnFailure();
                    int backoffMs = (int)Math.Min(8000, Math.Pow(2, attempt) * 250); // 250,500,1000,2000,4000,8000
                    await Task.Delay(backoffMs, ct);
                    attempt++;
                }
            }
            throw last ?? new Exception("network retry exceeded");
        }
    }
}