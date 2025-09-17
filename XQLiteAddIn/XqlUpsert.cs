// XqlUpsert.cs
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;

namespace XQLite.AddIn
{
    /// <summary>
    /// Excel 편집→GraphQL upsertRows 로 보내는 비동기 큐 + 내구성 아웃박스(NDJSON).
    /// - Enqueue(table,row) 로 행을 적재
    /// - 디바운스 타이머가 발사되면 테이블별 배치 업서트
    /// - 실패/오류는 outbox 로 기록 후, 매 회 Drain 끝에 재시도
    /// </summary>
    public static class XqlUpsert
    {
        // ---------------------------
        // 설정 & 상태
        // ---------------------------
        private static readonly ConcurrentQueue<(string table, Dictionary<string, object?> row)> _q = new();

        private static readonly object _timerGate = new();
        private static Timer? _timer;                 // 디바운스 타이머
        private static int _debounceMs = 2000;        // 디바운스 간격
        private static int _maxDegree = Math.Max(1, Environment.ProcessorCount / 2); // 테이블 병렬도(보수적)
        private static string _outboxPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "XQLite", "outbox.ndjson");

        private static readonly JsonSerializerOptions _jsonOpts = new()
        {
            WriteIndented = false
        };

        // ---------------------------
        // 초기화 & 설정
        // ---------------------------
        public static void Init(int debounceMs = 2000, string? outboxPath = null, int? maxParallelTables = null)
        {
            _debounceMs = Math.Max(200, debounceMs);
            if (!string.IsNullOrWhiteSpace(outboxPath))
                _outboxPath = outboxPath!;

            var dir = Path.GetDirectoryName(_outboxPath)!;
            Directory.CreateDirectory(dir);

            if (maxParallelTables is int md && md > 0) _maxDegree = md;
        }

        /// <summary>디바운스 간격 변경(밀리초)</summary>
        public static void SetDebounce(int ms)
        {
            _debounceMs = Math.Max(200, ms);
        }

        /// <summary>아웃박스 경로 변경(폴더 자동 생성)</summary>
        public static void SetOutboxPath(string path)
        {
            _outboxPath = path ?? _outboxPath;
            var dir = Path.GetDirectoryName(_outboxPath)!;
            Directory.CreateDirectory(dir);
        }

        // ---------------------------
        // API
        // ---------------------------
        /// <summary>
        /// 행을 전송 큐에 적재. 같은 테이블끼리 배치되어 전송됨.
        /// </summary>
        public static void Enqueue(string table, Dictionary<string, object?> row)
        {
            if (string.IsNullOrWhiteSpace(table) || row is null || row.Count == 0)
                return;

            _q.Enqueue((table, row));
            ArmTimer();
        }

        /// <summary>
        /// 대기열을 즉시 비움(전송 시도). 호출자 취소 토큰 제공 가능.
        /// </summary>
        public static Task FlushAsync(CancellationToken ct = default) => DrainAsync(ct);

        // ---------------------------
        // 내부: 디바운스 타이머
        // ---------------------------
        private static void ArmTimer()
        {
            lock (_timerGate)
            {
                _timer?.Dispose();
                // Timer 콜백의 async-void 위험을 피하기 위해 Task.Run 내부에서 호출
                _timer = new Timer(_ =>
                {
                    Task.Run(async () =>
                    {
                        try { await DrainAsync(); }
                        catch { /* no-throw */ }
                    });
                }, null, _debounceMs, Timeout.Infinite);
            }
        }

        // ---------------------------
        // 내부: Drain 파이프라인
        // ---------------------------
        private static async Task DrainAsync(CancellationToken ct = default)
        {
            // 1) 큐를 모두 빼서 테이블별로 모음
            var byTable = new Dictionary<string, List<Dictionary<string, object?>>>(StringComparer.OrdinalIgnoreCase);
            while (_q.TryDequeue(out var item))
            {
                if (!byTable.TryGetValue(item.table, out var list))
                    byTable[item.table] = list = new();
                list.Add(item.row);
            }
            if (byTable.Count == 0)
            {
                // 그래도 아웃박스 재시도는 해볼 가치 있음
                await TryReplayOutboxAsync(ct);
                return;
            }

            const string MUT = @"
mutation ($table:String!, $rows:[JSON!]!) {
  upsertRows(table:$table, rows:$rows) {
    affected
    max_row_version
    errors { code message }
  }
}";

            // 2) 테이블 단위 병렬 처리(보수적 병렬도)
            using var sem = new SemaphoreSlim(_maxDegree);
            var tasks = new List<Task>(byTable.Count);

            foreach (var (table, rows) in byTable)
            {
                await sem.WaitAsync(ct);
                tasks.Add(Task.Run(async () =>
                {
                    try
                    {
                        // 2-1) 동적 배치(행 수 + 대략 바이트 크기 기준)
                        int chunkSize = PickBatchSize(rows);

                        for (int i = 0; i < rows.Count; i += chunkSize)
                        {
                            ct.ThrowIfCancellationRequested();
                            var chunk = rows.Skip(i).Take(Math.Min(chunkSize, rows.Count - i)).ToList();
                            try
                            {
                                var resp = await XqlGraphQLClient.MutateAsync<UpsertResp>(MUT, new { table, rows = chunk }, ct);
                                var data = resp.Data?.upsertRows;

                                if (data?.errors is { Length: > 0 })
                                {
                                    XqlLog.Warn($"upsertRows errors={data.errors.Length} on {table}", table);
                                    await AppendOutboxAsync(table, chunk, ct);
                                }
                                else
                                {
                                    XqlLog.Info($"upsertRows ok affected={data?.affected ?? 0} {table}", table);
                                }
                            }
                            catch (Exception ex)
                            {
                                XqlLog.Warn($"upsertRows failed: {ex.Message} {table}", table);
                                await AppendOutboxAsync(table, chunk, ct);
                            }
                        }
                    }
                    finally
                    {
                        sem.Release();
                    }
                }, ct));
            }

            await Task.WhenAll(tasks);

            // 3) 끝나고 아웃박스 재시도
            await TryReplayOutboxAsync(ct);
        }

        // ---------------------------
        // 배치 전략
        // ---------------------------
        /// <summary>
        /// 대략적 바이트/행 수 기반으로 배치 크기 추정.
        /// - 목표 256KB 전후 / 최대 1500 행
        /// </summary>
        private static int PickBatchSize(List<Dictionary<string, object?>> rows)
        {
            int n = rows.Count;
            if (n <= 0) return 500;

            // 간단한 바이트 추정: key+value 문자열 길이
            static int RowSize(Dictionary<string, object?> r)
                => r.Sum(kv => (kv.Key?.Length ?? 0) + (kv.Value?.ToString()?.Length ?? 0));

            // 샘플 32개까지 평균
            var sample = rows.Take(Math.Min(32, n)).ToArray();
            double avg = Math.Max(16, sample.Average(RowSize)); // 너무 작게 나오면 최소 보정
            const int TargetBytes = 256 * 1024; // 256KB
            int byBytes = (int)Math.Clamp(TargetBytes / avg, 100, 1500);

            // 행수 절대 상한/하한도 적용
            if (n > 20000) return Math.Min(byBytes, 200);
            if (n > 5000) return Math.Min(byBytes, 1000);
            return Math.Min(byBytes, 1500);
        }

        // ---------------------------
        // 아웃박스 (내구성 큐)
        // ---------------------------
        private static async Task AppendOutboxAsync(string table, List<Dictionary<string, object?>> rows, CancellationToken ct = default)
        {
            try
            {
                using var fs = new FileStream(_outboxPath, FileMode.Append, FileAccess.Write, FileShare.Read);
                using var sw = new StreamWriter(fs, Encoding.UTF8);
                var now = DateTimeOffset.UtcNow;

                foreach (var r in rows)
                {
                    // NDJSON: { ts, table, row }
                    var line = JsonSerializer.Serialize(new OutItem
                    {
                        ts = now,
                        table = table,
                        row = r
                    }, _jsonOpts);
                    await sw.WriteLineAsync(line.AsMemory(), ct);
                }
            }
            catch (Exception ex)
            {
                XqlLog.Warn($"outbox append failed: {ex.Message}", table);
            }
        }

        private static async Task TryReplayOutboxAsync(CancellationToken ct = default)
        {
            if (!File.Exists(_outboxPath)) return;

            var tmp = _outboxPath + ".tmp";
            try { File.Move(_outboxPath, tmp, true); }
            catch
            {
                // 다른 프로세스가 사용 중일 수 있음
                return;
            }

            var byTable = new Dictionary<string, List<Dictionary<string, object?>>>(StringComparer.OrdinalIgnoreCase);
            foreach (var line in File.ReadLines(tmp))
            {
                try
                {
                    var obj = JsonSerializer.Deserialize<OutItem>(line);
                    if (obj?.row is null || string.IsNullOrWhiteSpace(obj.table)) continue;
                    if (!byTable.TryGetValue(obj.table, out var list)) byTable[obj.table] = list = new();
                    list.Add(obj.row);
                }
                catch { /* skip malformed line */ }
            }

            const string MUT = @"
mutation ($table:String!, $rows:[JSON!]!) {
  upsertRows(table:$table, rows:$rows) {
    affected
    max_row_version
    errors { code message }
  }
}";

            foreach (var (table, rows) in byTable)
            {
                int chunkSize = PickBatchSize(rows);
                for (int i = 0; i < rows.Count; i += chunkSize)
                {
                    ct.ThrowIfCancellationRequested();
                    var chunk = rows.Skip(i).Take(Math.Min(chunkSize, rows.Count - i)).ToList();
                    try
                    {
                        var resp = await XqlGraphQLClient.MutateAsync<UpsertResp>(MUT, new { table, rows = chunk }, ct);
                        var data = resp.Data?.upsertRows;
                        if (data?.errors is { Length: > 0 })
                        {
                            // 여전히 실패 → 원본 outbox 로 되돌림
                            await AppendOutboxAsync(table, chunk, ct);
                        }
                    }
                    catch
                    {
                        await AppendOutboxAsync(table, chunk, ct);
                    }
                }
            }

            try { File.Delete(tmp); } catch { /* best-effort */ }
        }

        // ---------------------------
        // 모델 (DTO)
        // ---------------------------
        private sealed class OutItem
        {
            public DateTimeOffset ts { get; set; }
            public string table { get; set; } = "";
            public Dictionary<string, object?>? row { get; set; }
        }

        /// <summary>upsertRows 응답 DTO</summary>
        public sealed class UpsertResp
        {
            public UpsertData? upsertRows { get; set; }
        }

        public sealed class UpsertData
        {
            public int affected { get; set; }
            public long max_row_version { get; set; }
            public Err[]? errors { get; set; }
        }

        public sealed class Err
        {
            public string code { get; set; } = "";
            public string message { get; set; } = "";
        }

        /// <summary>rows 쿼리 응답 DTO(스냅샷/익스포트에 사용)</summary>
        public sealed class RowsResp
        {
            public RowBlock[]? rows { get; set; }

            public sealed class RowBlock
            {
                public string table { get; set; } = string.Empty;
                public Dictionary<string, object?>[]? rows { get; set; }
                public long max_row_version { get; set; }
            }
        }
    }
}
