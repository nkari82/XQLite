using GraphQL;
using GraphQL.Client.Http;
using GraphQL.Client.Serializer.SystemTextJson;
using System;
#nullable enable
using System.Threading;
using System.Threading.Tasks;


namespace XQLite.AddIn;


public static class XqlGraphQLClient
{
    private static GraphQLHttpClient? _client;


    public static void Init(XqlConfig cfg)
    {
        _client = new GraphQLHttpClient(cfg.Endpoint, new SystemTextJsonSerializer());
        if (!string.IsNullOrEmpty(cfg.ApiKey)) _client.HttpClient.DefaultRequestHeaders.Add("x-api-key", cfg.ApiKey);
        if (!string.IsNullOrEmpty(cfg.Nickname)) _client.HttpClient.DefaultRequestHeaders.Add("x-actor", cfg.Nickname);
    }


    public static async Task<GraphQLResponse<T>> QueryAsync<T>(string query, object? vars = null, CancellationToken ct = default)
    {
        if (_client is null) throw new InvalidOperationException("GraphQL client not initialized");
        var req = new GraphQLRequest { Query = query, Variables = vars };
        return await _client.SendQueryAsync<T>(req, ct);
    }


    public static async Task<GraphQLResponse<T>> MutateAsync<T>(string query, object? vars = null, CancellationToken ct = default)
    {
        if (_client is null) throw new InvalidOperationException("GraphQL client not initialized");
        var req = new GraphQLRequest { Query = query, Variables = vars };
        return await _client.SendMutationAsync<T>(req, ct);
    }
}