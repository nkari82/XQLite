#nullable enable
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Text.Json.Serialization;
using System.Threading.Tasks;

namespace XQLite.AddIn;

public static class XqlSyncService
{
    // 업서트 버퍼: (table, row)
    private static readonly ConcurrentQueue<(string table, Dictionary<string, object?> row)> _q = new();
    private static System.Timers.Timer? _debounce;
    private static System.Timers.Timer? _pull;
    private static XqlConfig? _cfg;
    private static long _sinceVersion = 0;

    public static void Start(XqlConfig cfg)
    {
        _cfg = cfg;
        _debounce = new System.Timers.Timer(Math.Max(500, cfg.DebounceMs)) { AutoReset = false };
        _debounce.Elapsed += async (_, __) => await FlushAsync();

        _pull = new System.Timers.Timer(Math.Max(1000, cfg.PullSec * 1000)) { AutoReset = true };
        _pull.Elapsed += async (_, __) => await PullAsync();

        _pull.Start();
    }

    public static void Stop()
    {
        if (_debounce is not null) { _debounce.Stop(); _debounce.Dispose(); _debounce = null; }
        if (_pull is not null) { _pull.Stop(); _pull.Dispose(); _pull = null; }
        _cfg = null;
    }

    // 외부(다음 스텝의 시트 이벤트 등)에서 행을 큐잉하는 API
    public static void QueueUpsert(string table, Dictionary<string, object?> row)
    {
        _q.Enqueue((table, row));
        _debounce?.Stop();
        if (_cfg is not null) _debounce!.Interval = Math.Max(500, _cfg.DebounceMs);
        _debounce?.Start();
    }

    private static async Task FlushAsync()
    {
        if (_q.IsEmpty) return;

        // 테이블별로 묶기
        var buckets = new Dictionary<string, List<Dictionary<string, object?>>>(StringComparer.OrdinalIgnoreCase);
        while (_q.TryDequeue(out var item))
        {
            if (!buckets.TryGetValue(item.table, out var list))
            { list = new List<Dictionary<string, object?>>(); buckets[item.table] = list; }
            list.Add(item.row);
        }

        const string m = @"mutation ($table:String!,$rows:[JSON!]!){
  upsertRows(table:$table, rows:$rows){ affected, conflicts, errors{code,message,path}, max_row_version }
}";
        foreach (var kv in buckets)
        {
            try
            {
                var resp = await XqlGraphQLClient.MutateAsync<UpsertResp>(m, new { table = kv.Key, rows = kv.Value });
                var data = resp.Data?.upsertRows;
                if (data is not null)
                {
                    _sinceVersion = Math.Max(_sinceVersion, data.max_row_version);
                    // 에러/컨플릭트는 지금 스텝에선 로깅만 (시트 표시 없음)
                    if (data.errors is not null && data.errors.Length > 0)
                    {
                        System.Diagnostics.Debug.WriteLine($"upsertRows errors: {data.errors.Length}");
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"upsertRows failed: {ex.Message}");
            }
        }
    }

    private static async Task PullAsync()
    {
        const string q = "query($since:Long){ rows(since_version:$since){ table, rows, max_row_version } }";
        try
        {
            var resp = await XqlGraphQLClient.QueryAsync<RowsResp>(q, new { since = _sinceVersion });
            var blocks = resp.Data?.rows;
            if (blocks is null) return;

            foreach (var blk in blocks)
            {
                _sinceVersion = Math.Max(_sinceVersion, blk.max_row_version);
                // STEP 2: 시트 반영 없음. 여기서는 단순히 최신 row_version만 따라간다.
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"pull failed: {ex.Message}");
        }
    }

    // ── 응답 DTO (dynamic보다 안전하게 최소 타입 정의)
    public sealed class UpsertResp { public UpsertPayload? upsertRows { get; set; } }
    public sealed class UpsertPayload
    {
        public int affected { get; set; }
        public ErrorDto[]? errors { get; set; }
        public ConflictDto[]? conflicts { get; set; }
        public long max_row_version { get; set; }
    }
    public sealed class ErrorDto { public string? code { get; set; } public string? message { get; set; } public string? path { get; set; } }
    public sealed class ConflictDto { public string? key { get; set; } public string? reason { get; set; } }

    public sealed class RowsResp { public RowBlock[]? rows { get; set; } }
    public sealed class RowBlock
    {
        public string table { get; set; } = string.Empty;
        public Dictionary<string, object?>[]? rows { get; set; }
        public long max_row_version { get; set; }
    }
}