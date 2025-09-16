// ==============================================
// XQLite C# Port — STEP 12
// 로컬 Mock GraphQL 서버 + 통합 테스트 하네스
// • 목표: 실제 서버 없이도 Add-in 전체 플로우(업서트/풀/구독/프레즌스) 개발·디버그 가능
// • Tech: .NET 8 Minimal API + GraphQL.NET + WebSocket Subscriptions
// • Endpoints: HTTP(s) /graphql, WS(s) /graphql  (Same path; transport에 따라 라우팅)
// • Schema: rows, upsertRows, presenceHeartbeat, presence, rowsChanged(subscription)
// • 저장소: InMemory (Dictionary<string,List<Dictionary<string,object?>>>) + row_version 증가
// ==============================================

// ─────────────────────────────
// Project: Xql.MockServer (Console)
// csproj (요지)
// ─────────────────────────────
/*
<Project Sdk="Microsoft.NET.Sdk">
  <PropertyGroup>
    <OutputType>Exe</OutputType>
    <TargetFramework>net8.0</TargetFramework>
    <Nullable>enable</Nullable>
    <ImplicitUsings>enable</ImplicitUsings>
  </PropertyGroup>
  <ItemGroup>
    <PackageReference Include="GraphQL" Version="7.*" />
    <PackageReference Include="GraphQL.SystemTextJson" Version="7.*" />
    <PackageReference Include="GraphQL.Server.Transports.AspNetCore" Version="7.*" />
    <PackageReference Include="GraphQL.Server.Transports.WebSockets" Version="7.*" />
    <PackageReference Include="GraphQL.Server.Ui.Playground" Version="7.*" />
  </ItemGroup>
</Project>
*/

// ─────────────────────────────
// File: Program.cs (Xql.MockServer)
// ─────────────────────────────
using GraphQL;
using GraphQL.Federation.Types;
using GraphQL.Resolvers;
using GraphQL.Server;
using GraphQL.Server.Transports.AspNetCore;
using GraphQL.Server.Transports.WebSockets;
using GraphQL.Server.Ui.Playground;
using GraphQL.SystemTextJson;
using GraphQL.Types;
using Microsoft.AspNetCore.Builder;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using System.Collections.Concurrent;

var builder = WebApplication.CreateBuilder(args);

// GraphQL DI
builder.Services.AddSingleton<XqlStore>();
builder.Services.AddSingleton<XqlQuery>();
builder.Services.AddSingleton<XqlMutation>();
builder.Services.AddSingleton<XqlSubscription>();
builder.Services.AddSingleton<XqlSchema>();

builder.Services.AddGraphQL(o =>
{
    o.EnableMetrics = false;
}).AddSystemTextJson();

builder.Services.AddGraphQLWebSockets().AddGraphQLHttpTransport();

var app = builder.Build();
app.UseWebSockets();
app.UseRouting();

app.UseEndpoints(endpoints =>
{
    endpoints.MapGraphQL("/graphql");
    endpoints.MapGraphQLWebSockets("/graphql");
});

app.UseGraphQLPlayground(new PlaygroundOptions
{
    GraphQLEndPoint = "/graphql",
    SubscriptionsEndPoint = "/graphql"
});

app.Run();

// ─────────────────────────────
// File: XqlStore.cs — InMemory 저장 + PubSub
// ─────────────────────────────
public sealed class XqlStore
{
    public long MaxRowVersion { get; private set; }
    private readonly ConcurrentDictionary<string, List<Dictionary<string, object?>>> _tables = new(StringComparer.OrdinalIgnoreCase);
    private readonly ConcurrentDictionary<string, DateTime> _presence = new(StringComparer.OrdinalIgnoreCase);

    // 간단한 pubsub — 구독자에게 푸시
    private readonly IObservable<RowBlock[]> _stream;
    private readonly IObserver<RowBlock[]> _observer;

    public XqlStore()
    {
        var subj = new SimpleSubject<RowBlock[]>();
        _stream = subj; _observer = subj;
    }

    public (List<RowBlock> blocks, long max) GetRows(long since)
    {
        // since 무시하고 전 테이블 스냅 (mock 간단화)
        var list = new List<RowBlock>();
        foreach (var kv in _tables)
        {
            list.Add(new RowBlock
            {
                table = kv.Key,
                rows = kv.Value.ToArray(),
                max_row_version = MaxRowVersion
            });
        }
        return (list, MaxRowVersion);
    }

    public UpsertPayload Upsert(string table, List<Dictionary<string, object?>> rows)
    {
        var list = _tables.GetOrAdd(table, _ => new());
        // 키(id) 기준 단순 merge (id 없으면 append)
        int affected = 0;
        foreach (var row in rows)
        {
            string? id = null;
            if (row.TryGetValue("id", out var v) && v != null) id = Convert.ToString(v);
            var found = id == null ? null : list.FirstOrDefault(r => Convert.ToString(r.GetValueOrDefault("id")) == id);
            if (found != null)
            {
                foreach (var k in row.Keys) found[k] = row[k];
            }
            else { list.Add(new Dictionary<string, object?>(row, StringComparer.OrdinalIgnoreCase)); }
            affected++;
        }
        MaxRowVersion++;
        var block = new RowBlock { table = table, rows = rows.ToArray(), max_row_version = MaxRowVersion };
        _observer.OnNext(new[] { block });
        return new UpsertPayload { affected = affected, errors = Array.Empty<ErrorDto>(), conflicts = Array.Empty<ConflictDto>(), max_row_version = MaxRowVersion };
    }

    public void Heartbeat(string nickname, string? sheet, string? cell)
    {
        _presence[nickname] = DateTime.UtcNow;
    }

    public PresenceItem[] Presence()
    {
        var now = DateTime.UtcNow;
        return _presence.Select(kv => new PresenceItem { nickname = kv.Key, sheet = "-", cell = "-", updated_at = kv.Value.ToString("o") }).ToArray();
    }

    public IObservable<RowBlock[]> RowsChangedStream() => _stream;
}

public sealed class SimpleSubject<T> : IObservable<T>, IObserver<T>
{
    private readonly List<IObserver<T>> _observers = new();
    public IDisposable Subscribe(IObserver<T> observer) { _observers.Add(observer); return new Unsub(_observers, observer); }
    public void OnCompleted() { foreach (var o in _observers) o.OnCompleted(); }
    public void OnError(Exception error) { foreach (var o in _observers) o.OnError(error); }
    public void OnNext(T value) { foreach (var o in _observers) o.OnNext(value); }
    private sealed class Unsub : IDisposable { private readonly List<IObserver<T>> _list; private readonly IObserver<T> _obs; public Unsub(List<IObserver<T>> l, IObserver<T> o) { _list = l; _obs = o; } public void Dispose() { _list.Remove(_obs); } }
}

// ─────────────────────────────
// File: GraphQL 타입들
// ─────────────────────────────
public sealed class RowBlock
{
    public string table { get; set; } = string.Empty;
    public Dictionary<string, object?>[] rows { get; set; } = Array.Empty<Dictionary<string, object?>>();
    public long max_row_version { get; set; }
}
public sealed class UpsertPayload
{
    public int affected { get; set; }
    public ErrorDto[] errors { get; set; } = Array.Empty<ErrorDto>();
    public ConflictDto[] conflicts { get; set; } = Array.Empty<ConflictDto>();
    public long max_row_version { get; set; }
}
public sealed class ErrorDto { public string code { get; set; } = "E_GENERIC"; public string message { get; set; } = ""; public string? path { get; set; } }
public sealed class ConflictDto { public string? key { get; set; } public string? reason { get; set; } }
public sealed class PresenceItem { public string? nickname { get; set; } public string? sheet { get; set; } public string? cell { get; set; } public string? updated_at { get; set; } }

// GraphQL.NET ObjectGraphType들
public sealed class RowBlockType : ObjectGraphType<RowBlock>
{
    public RowBlockType() { Field(x => x.table); Field(x => x.rows, type: typeof(ListGraphType<AnyScalarGraphType>)); Field(x => x.max_row_version); }
}
public sealed class UpsertPayloadType : ObjectGraphType<UpsertPayload>
{
    public UpsertPayloadType() { Field(x => x.affected); Field<ListGraphType<ErrorType>>("errors"); Field<ListGraphType<ConflictType>>("conflicts"); Field(x => x.max_row_version); }
}
public sealed class ErrorType : ObjectGraphType<ErrorDto> { public ErrorType() { Field(x => x.code); Field(x => x.message); Field(x => x.path, nullable: true); } }
public sealed class ConflictType : ObjectGraphType<ConflictDto> { public ConflictType() { Field(x => x.key, nullable: true); Field(x => x.reason, nullable: true); } }
public sealed class PresenceType : ObjectGraphType<PresenceItem> { public PresenceType() { Field(x => x.nickname, nullable: true); Field(x => x.sheet, nullable: true); Field(x => x.cell, nullable: true); Field(x => x.updated_at, nullable: true); } }

public sealed class XqlQuery : ObjectGraphType
{
    public XqlQuery(XqlStore store)
    {
        Field<ListGraphType<RowBlockType>>("rows")
            .Argument<LongGraphType>("since_version")
            .Resolve(ctx => store.GetRows(ctx.GetArgument<long>("since_version", 0)).blocks);

        Field<ListGraphType<PresenceType>>("presence")
            .Resolve(_ => store.Presence());

        Field<StringGraphType>("health").Resolve(_ => "ok");
    }
}

public sealed class XqlMutation : ObjectGraphType
{
    public XqlMutation(XqlStore store)
    {
        Field<UpsertPayloadType>("upsertRows")
            .Argument<NonNullGraphType<StringGraphType>>("table")
            .Argument<NonNullGraphType<ListGraphType<AnyScalarGraphType>>>("rows")
            .Resolve(ctx =>
            {
                var table = ctx.GetArgument<string>("table");
                var rows = ctx.GetArgument<List<Dictionary<string, object?>>>("rows");
                return store.Upsert(table, rows);
            });

        Field<BooleanGraphType>("presenceHeartbeat")
            .Argument<NonNullGraphType<StringGraphType>>("nickname")
            .Resolve(ctx => { store.Heartbeat(ctx.GetArgument<string>("nickname"), null, null); return true; });
    }
}

public sealed class XqlSubscription : ObjectGraphType
{
    public XqlSubscription(XqlStore store)
    {
        AddField(new EventStreamFieldType
        {
            Name = "rowsChanged",
            Arguments = new QueryArguments(new QueryArgument<LongGraphType> { Name = "since_version" }),
            Type = typeof(ListGraphType<RowBlockType>),
            Resolver = new FuncFieldResolver<List<RowBlock>>(ctx => ctx.Source as List<RowBlock> ?? new()),
            Subscriber = new EventStreamResolver<List<RowBlock>>(ctx => store.RowsChangedStream())
        });
    }
}

public sealed class XqlSchema : Schema
{
    public XqlSchema(IServiceProvider sp) : base(sp)
    {
        Query = sp.GetRequiredService<XqlQuery>();
        Mutation = sp.GetRequiredService<XqlMutation>();
        Subscription = sp.GetRequiredService<XqlSubscription>();
        RegisterType<AnyScalarGraphType>();
    }
}

// ─────────────────────────────
// 사용법
// ─────────────────────────────
/*
1) 솔루션에 "Xql.MockServer" 콘솔 프로젝트 추가 후 위 파일들 생성 → 빌드/실행
   - 기본 URL: http://localhost:5000/graphql (Playground 포함)
   - WS: ws://localhost:5000/graphql
2) Add-in Config에서 Endpoint를 해당 URL로 설정
3) 기능 확인
   - Excel에서 테이블 셀 수정 → upsertRows 호출 → MockServer 메모리 저장소 반영
   - Inspector에서 로그 확인, Export로 rows 스냅샷 저장
   - (STEP 11에서 붙인) Subscription 모드면 rowsChanged 실시간 수신 확인
4) 저장 지속성 필요 시: XqlStore에 파일 스냅샷(Load/Save) 몇 줄 추가하면 됨(JSON 직렬화)
*/
