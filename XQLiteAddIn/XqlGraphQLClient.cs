// XqlGraphQLClient.cs (추가/수정)
using GraphQL;
using GraphQL.Client.Http;
using GraphQL.Client.Serializer.SystemTextJson;
using System;
using System.Net.Http.Headers;
using System.Threading;
using System.Threading.Tasks;

namespace XQLite.AddIn
{
    public static class XqlGraphQLClient
    {
        private static GraphQLHttpClient? _client;
        private static string _endpoint = "";

        public static void Init(XqlConfig cfg)
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

            _client = new GraphQLHttpClient(options, new SystemTextJsonSerializer());

            string apiKey = cfg.ApiKey == "__DPAPI__" ? XqlSecrets.LoadApiKey() : cfg.ApiKey;
            if (!string.IsNullOrEmpty(apiKey))
                _client.HttpClient.DefaultRequestHeaders.Add("x-api-key", apiKey);
            if (!string.IsNullOrEmpty(cfg.Nickname))
                _client.HttpClient.DefaultRequestHeaders.Add("x-actor", cfg.Nickname);
            _client.HttpClient.DefaultRequestHeaders.UserAgent.Add(new ProductInfoHeaderValue("XQLite.AddIn", "1.0"));
        }

        public static Task<GraphQLResponse<T>> QueryAsync<T>(string q, object? vars = null, CancellationToken ct = default)
          => XqlNetPolicy.WithRetryAsync(() => _client!.SendQueryAsync<T>(new GraphQLRequest { Query = q, Variables = vars }, ct), ct);

        public static Task<GraphQLResponse<T>> MutateAsync<T>(string q, object? vars = null, CancellationToken ct = default)
          => XqlNetPolicy.WithRetryAsync(() => _client!.SendMutationAsync<T>(new GraphQLRequest { Query = q, Variables = vars }, ct), ct);

        // ★ 구독 스트림
        public static IObservable<GraphQLResponse<T>> Subscribe<T>(string sub, object? vars = null)
        {
            var req = new GraphQLRequest { Query = sub, Variables = vars };
            return _client!.CreateSubscriptionStream<T>(req);
        }
    }
}