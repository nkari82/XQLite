// XqlGraphQLClient.cs (추가/수정)
using GraphQL;
using GraphQL.Client.Http;
using GraphQL.Client.Serializer.Newtonsoft;
using System;
using System.Net.Http.Headers;
using System.Threading;
using System.Threading.Tasks;

namespace XQLite.AddIn
{
    internal static class XqlGraphQLClient
    {
        internal static class NetPolicy
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


        private static GraphQLHttpClient? _client;
        private static string _endpoint = "";

        internal static void Init(XqlConfig cfg)
        {
            _endpoint = cfg.Endpoint.Trim();
            var http = new Uri(_endpoint);
            var ws = new Uri((http.Scheme == "https" ? "wss" : "ws") + "://" + http.Host + (http.IsDefaultPort ? "" : ":" + http.Port) + http.PathAndQuery);

            var options = new GraphQLHttpClientOptions
            {
                EndPoint = http,
                WebSocketEndPoint = ws,
                // 쿼리/뮤테이션은 HTTP, 구독은 WebSocket 사용(기본값)
            };

            _client = new GraphQLHttpClient(options, new NewtonsoftJsonSerializer());

            string apiKey = cfg.ApiKey == "__DPAPI__" ? XqlSecrets.LoadApiKey() : cfg.ApiKey;
            if (!string.IsNullOrEmpty(apiKey))
                _client.HttpClient.DefaultRequestHeaders.Add("x-api-key", apiKey);
            if (!string.IsNullOrEmpty(cfg.Nickname))
                _client.HttpClient.DefaultRequestHeaders.Add("x-actor", cfg.Nickname);
            _client.HttpClient.DefaultRequestHeaders.UserAgent.Add(new ProductInfoHeaderValue("XQLite.AddIn", "1.0"));
        }

        internal static Task<GraphQLResponse<T>> QueryAsync<T>(string q, object? vars = null, CancellationToken ct = default)
          => NetPolicy.WithRetryAsync(() => _client!.SendQueryAsync<T>(new GraphQLRequest { Query = q, Variables = vars }, ct), ct);

        internal static Task<GraphQLResponse<T>> MutateAsync<T>(string q, object? vars = null, CancellationToken ct = default)
          => NetPolicy.WithRetryAsync(() => _client!.SendMutationAsync<T>(new GraphQLRequest { Query = q, Variables = vars }, ct), ct);

        // 구독 스트림
        internal static IObservable<GraphQLResponse<T>> Subscribe<T>(string sub, object? vars = null)
        {
            var req = new GraphQLRequest { Query = sub, Variables = vars };
            return _client!.CreateSubscriptionStream<T>(req);
        }

        internal static void Dispose()
        {
            try { _client?.Dispose(); } catch { }
            _client = null;
        }
    }
}