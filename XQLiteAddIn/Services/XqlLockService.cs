using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace XQLite.AddIn
{
    internal static class XqlLockService
    {
        internal sealed class LockInfo
        {
            public string lockId { get; set; } = string.Empty;
            public string owner { get; set; } = string.Empty;
            public string resource { get; set; } = string.Empty; // e.g., table/items or sheet/Design
            public Dictionary<string, object?> scope { get; set; } = new(); // type,column,rowId,address...
            public DateTime expiresAt { get; set; }
        }

        public sealed class AcquireResp { public AcquireData? acquireLock { get; set; } }
        public sealed class AcquireData { public bool ok { get; set; } public string? lockId { get; set; } public int ttlSec { get; set; } }
        public sealed class ReleaseResp { public ReleaseData? releaseLock { get; set; } }
        public sealed class ReleaseData { public bool ok { get; set; } }
        public sealed class LocksResp { public LockInfo[]? locks { get; set; } }

        private static readonly ConcurrentDictionary<string, LockInfo> _locksById = new();
        private static readonly ConcurrentDictionary<string, LockInfo> _locksByKey = new(StringComparer.OrdinalIgnoreCase);
        private static Timer? _timer; // 3s 주기로 refresh

        internal static void Start()
        {
            _timer = new Timer(async _ => await RefreshAsync(), null, 0, 3000);
        }
        internal static void Stop() { _timer?.Dispose(); _timer = null; _locksById.Clear(); _locksByKey.Clear(); }

        internal static async Task<bool> AcquireColumnAsync(string table, string column, int ttlSec = 10)
            => await AcquireAsync(new() { { "type", "column" }, { "table", table }, { "column", column } }, ttlSec);

        internal static async Task<bool> AcquireCellAsync(string sheet, string address, int ttlSec = 10)
            => await AcquireAsync(new() { { "type", "cell" }, { "sheet", sheet }, { "address", address } }, ttlSec);

        internal static async Task<bool> ReleaseAsync(string lockId)
        {
            try
            {
                const string m = "mutation($id:String!){ releaseLock(lockId:$id){ ok } }";
                var r = await XqlGraphQLClient.MutateAsync<ReleaseResp>(m, new { id = lockId });
                return r.Data?.releaseLock?.ok == true;
            }
            catch { return false; }
            finally
            {
                _locksById.TryRemove(lockId, out _);
                foreach (var kv in _locksByKey.Where(kv => kv.Value.lockId == lockId).ToList()) _locksByKey.TryRemove(kv.Key, out _);
            }
        }

        private static async Task<bool> AcquireAsync(Dictionary<string, object?> scope, int ttlSec)
        {
            try
            {
                var resKey = MakeKey(scope);
                const string m = "mutation($res:String!,$scope:JSON!,$ttl:Int!){ acquireLock(resource:$res, scope:$scope, ttlSec:$ttl){ ok, lockId, ttlSec } }";
                var r = await XqlGraphQLClient.MutateAsync<AcquireResp>(m, new { res = resKey, scope, ttl = ttlSec });
                var data = r.Data?.acquireLock; if (data?.ok != true || string.IsNullOrEmpty(data.lockId)) return false;
                var info = new LockInfo { lockId = data.lockId!, owner = XqlAddIn.Cfg?.Nickname ?? Environment.UserName, resource = resKey, scope = scope, expiresAt = DateTime.UtcNow.AddSeconds(data.ttlSec) };
                _locksById[info.lockId] = info; _locksByKey[resKey] = info; return true;
            }
            catch { return false; }
        }

        private static async Task RefreshAsync()
        {
            try
            {
                const string q = "query{ locks{ lockId owner resource scope expiresAt } }";
                var r = await XqlGraphQLClient.QueryAsync<LocksResp>(q);
                var list = r.Data?.locks ?? Array.Empty<LockInfo>();
                var now = DateTime.UtcNow;
                _locksById.Clear(); _locksByKey.Clear();
                foreach (var l in list) { _locksById[l.lockId] = l; _locksByKey[l.resource] = l; }
            }
            catch { /* 서버가 락을 지원 안하면 조용히 무시 */ }
        }

        internal static bool IsLockedColumn(string table, string column)
        {
            var key = MakeKey(new() { { "type", "column" }, { "table", table }, { "column", column } });
            return _locksByKey.ContainsKey(key);
        }

        internal static bool IsLockedCell(string sheet, string address)
        {
            var key = MakeKey(new() { { "type", "cell" }, { "sheet", sheet }, { "address", address } });
            return _locksByKey.ContainsKey(key);
        }

        private static string MakeKey(Dictionary<string, object?> scope)
        {
            // 정규화된 리소스 키: type=column => "column:table/col" ; type=cell => "cell:sheet/address"
            var type = Convert.ToString(scope.GetValueOrDefault("type")) ?? "";
            if (string.Equals(type, "column", StringComparison.OrdinalIgnoreCase))
            {
                var t = Convert.ToString(scope.GetValueOrDefault("table")) ?? "";
                var c = Convert.ToString(scope.GetValueOrDefault("column")) ?? "";
                return $"column:{t}/{c}";
            }
            if (string.Equals(type, "cell", StringComparison.OrdinalIgnoreCase))
            {
                var s = Convert.ToString(scope.GetValueOrDefault("sheet")) ?? "";
                var a = Convert.ToString(scope.GetValueOrDefault("address")) ?? "";
                return $"cell:{s}/{a.ToUpperInvariant()}";
            }
            // fallback
            return "misc:" + string.Join(";", scope.Select(kv => kv.Key + "=" + kv.Value));
        }
    }
}