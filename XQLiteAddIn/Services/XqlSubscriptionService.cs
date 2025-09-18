// XqlSubscriptionService.cs (신규)
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace XQLite.AddIn
{
    internal static class XqlSubscriptionService
    {
        private static IDisposable? _sub;
        private static long _sinceVersion;

        internal static void Start(long startSince = 0)
        {
            _sinceVersion = startSince;

            SubscribeRowsChanged();
        }

        internal static void Stop()
        {
            try { _sub?.Dispose(); } catch { }
            _sub = null;
        }

        private static void SubscribeRowsChanged()
        {
            // 서버 구독 예시(스키마에 맞게 조정):
            // subscription($since: Long){ rowsChanged(since_version:$since){ table, rows, max_row_version } }
            const string SUB = "subscription($since: Long){ rowsChanged(since_version:$since){ table rows max_row_version } }";
            var stream = XqlGraphQLClient.Subscribe<RowsPayload>(SUB, new { since = _sinceVersion });

            _sub = stream.Subscribe(
                onNext: resp =>
                {
                    var data = resp.Data?.rowsChanged;
                    if (data == null) return;

                    // 당장 시트는 건드리지 않고, 버전만 캐치업 + 로그만
                    foreach (var block in data)
                        _sinceVersion = Math.Max(_sinceVersion, block.max_row_version);

                    XqlLog.Info($"sub rowsChanged since={_sinceVersion}");
                },
                onError: ex =>
                {
                    XqlLog.Warn($"sub error: {ex.Message}");
                    // 재시도 간단 정책: 약간의 백오프 후 재구독
                    Task.Run(async () =>
                    {
                        await Task.Delay(2000);
                        SubscribeRowsChanged();
                    });
                },
                onCompleted: () =>
                {
                    XqlLog.Warn("sub completed");
                    // 필요시 재구독
                    Task.Run(async () =>
                    {
                        await Task.Delay(2000);
                        SubscribeRowsChanged();
                    });
                }
            );
        }

        // 서버 응답 DTO
        internal sealed class RowsPayload
        {
            internal RowBlock[]? rowsChanged { get; set; }
        }

        internal sealed class RowBlock
        {
            internal string table { get; set; } = "";
            internal Dictionary<string, object?>[]? rows { get; set; }
            internal long max_row_version { get; set; }
        }
    }
}