using System;
using System.Diagnostics;

namespace XQLite.AddIn
{
    public static class XqlPerf
    {
        public static (IDisposable scope, Action<long> done) Scope(string name, string table = "*")
        {
            var sw = Stopwatch.StartNew();
            return (new ScopeDisposable(() =>
            {
                var ms = sw.ElapsedMilliseconds; XqlLog.Info($"{name} took {ms} ms");
            }), bytes =>
            {
                var ms = Math.Max(1, sw.ElapsedMilliseconds);
                var kbps = (bytes / 1024.0) / (ms / 1000.0);
                XqlLog.Info($"{name}: ~{kbps:F1} KB/s (approx) on {table}");
            }
            );
        }

        private sealed class ScopeDisposable : IDisposable
        { private readonly Action _on; public ScopeDisposable(Action on) { _on = on; } public void Dispose() { _on(); } }
    }
}