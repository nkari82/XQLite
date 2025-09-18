// XqlFileLogger.cs (추가/보강)
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading;

namespace XQLite.AddIn
{
    internal static class XqlFileLogger
    {
        internal sealed class LogItem
        {
            public DateTime At { get; init; }
            public string Level { get; init; } = "";
            public string Table { get; init; } = "";
            public string Message { get; init; } = "";
            public string[]? Details { get; init; }
        }

        private static readonly BlockingCollection<string> _q = new(new ConcurrentQueue<string>());
        private static Thread? _worker;
        private static volatile bool _running;
        private static readonly object _fileGate = new();
        private static string _logDir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "XQLite", "logs");
        private static string _logPath = Path.Combine(_logDir, "xqlite.log");

        // --- 메모리 버퍼(최근 로그) & 드레인 지원 ---
        private static readonly ConcurrentQueue<LogItem> _recent = new();
        private static int _recentCapacity = 2000;

        internal static void Start(string? dir = null, int recentCapacity = 2000)
        {
            if (dir != null) { _logDir = dir; _logPath = Path.Combine(_logDir, "xqlite.log"); }
            _recentCapacity = Math.Max(100, recentCapacity);
            Directory.CreateDirectory(_logDir);
            _running = true;
            _worker ??= new Thread(WriterLoop) { IsBackground = true, Name = "XqlFileLogger" };
            if (!_worker.IsAlive) _worker.Start();
        }

        internal static void Stop()
        {
            _running = false;
            _q.CompleteAdding();
            try { _worker?.Join(1500); } catch { }
            _worker = null;
        }

        public static void Write(string level, string table, string message, params string[] details)
        {
            var now = DateTime.Now;
            var detailStr = (details != null && details.Length > 0) ? (" | " + string.Join(" | ", details)) : "";
            var line = $"{now:yyyy-MM-dd HH:mm:ss.fff}\t{level}\t{table}\t{message}{detailStr}";
            _q.Add(line);

            // 메모리 버퍼에도 동일 아이템 적재
            _recent.Enqueue(new LogItem
            {
                At = now,
                Level = level,
                Table = table ?? "",
                Message = message ?? "",
                Details = (details != null && details.Length > 0) ? details : null
            });
            // 용량 유지
            while (_recent.Count > _recentCapacity) _recent.TryDequeue(out _);
        }

        public static void Info(string table, string message, params string[] details) => Write("INFO", table, message, details);
        public static void Warn(string table, string message, params string[] details) => Write("WARN", table, message, details);
        public static void Error(string table, string message, params string[] details) => Write("ERR", table, message, details);

        // ✅ 여기 추가: 최근 로그를 꺼내오는 드레인 메서드 (XqlSyncService.TakeLogs 대체)
        public static IReadOnlyList<LogItem> TakeLogs(int max = 200)
        {
            if (max <= 0) return Array.Empty<LogItem>();
            var list = new List<LogItem>(Math.Min(max, _recent.Count));
            for (int i = 0; i < max && _recent.TryDequeue(out var item); i++)
                list.Add(item);
            return list;
        }

        private static void WriterLoop()
        {
            try
            {
                using var fs = new FileStream(_logPath, FileMode.Append, FileAccess.Write, FileShare.Read);
                using var sw = new StreamWriter(fs, Encoding.UTF8) { AutoFlush = true };
                foreach (var line in _q.GetConsumingEnumerable())
                {
                    if (!_running) break;
                    lock (_fileGate) { sw.WriteLine(line); }
                }
            }
            catch { /* 파일 쓰기 실패는 무시(버퍼에만 남음) */ }
        }
    }
}
