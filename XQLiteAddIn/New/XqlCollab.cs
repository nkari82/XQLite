// XqlCollab.cs
using System;
using System.Collections.Concurrent;
using System.Threading;

using GraphQL;
using GraphQL.Client.Http;
using GraphQL.Client.Serializer.Newtonsoft;
using Newtonsoft.Json.Linq;

namespace XQLite.AddIn
{
    /// <summary>
    /// 실시간 협업(프레즌스 + 시트/셀 락)
    /// - Heartbeat: 2.5초 주기, TTL: 10초
    /// - 로컬 사전으로 즉시 표시/판단하고, 가능하면 서버에 동기화
    /// </summary>
    internal sealed class XqlCollab : IDisposable
    {
        private readonly ConcurrentDictionary<string, DateTime> _presence = new(StringComparer.Ordinal);
        private readonly ConcurrentDictionary<string, string> _cellLocks = new(StringComparer.Ordinal); // key: "Sheet!A1", val: nickname

        private readonly Timer _ttlSweep;
        private readonly Timer _hbTimer;

        private readonly Backend? _backend; // 서버 연동은 선택

        private readonly int _ttlSec;
        private readonly int _hbMs;

        public XqlCollab(string? endpoint = null, string? apiKey = null, int ttlSeconds = 10, int heartbeatMs = 2500)
        {
            _ttlSec = Math.Max(5, ttlSeconds);
            _hbMs = Math.Max(1000, heartbeatMs);

            if (!string.IsNullOrWhiteSpace(endpoint))
                _backend = new Backend(new Uri(endpoint!), apiKey ?? "");

            _ttlSweep = new Timer(_ => Sweep(), null, 5000, 2000);
            _hbTimer = new Timer(_ => SendPeriodicHeartbeat(), null, Timeout.Infinite, Timeout.Infinite);
        }

        public void Dispose()
        {
            try { _ttlSweep.Dispose(); } catch { }
            try { _hbTimer.Dispose(); } catch { }
            try { _backend?.Dispose(); } catch { }
        }

        // ===== Presence =====
        private volatile string _lastNickname = "anonymous";
        private volatile string _lastCellRef = "";

        /// <summary>UI에서 선택 변경 시 호출: 현재 선택 셀을 알려주면 디바운스된 하트비트로 반영됨.</summary>
        public void NotifySelection(string nickname, string sheetExAddr)
        {
            _lastNickname = string.IsNullOrWhiteSpace(nickname) ? "anonymous" : nickname;
            _lastCellRef = sheetExAddr ?? "";
            // 2.5초 디바운스 하트비트
            _hbTimer.Change(_hbMs, Timeout.Infinite);
        }

        /// <summary>즉시 하트비트(메뉴/버튼 등에서 호출 가능)</summary>
        public void Heartbeat(string nickname, string? cellRef = null)
        {
            try
            {
                nickname = string.IsNullOrWhiteSpace(nickname) ? "anonymous" : nickname;
                var cref = cellRef ?? _lastCellRef;

                _presence[nickname] = DateTime.UtcNow;
                if (!string.IsNullOrEmpty(cref))
                    _cellLocks.TryAdd(cref, nickname);

                // 서버 동기화(선택)
                _backend?.PresenceHeartbeat(nickname, cref);
            }
            catch { }
        }

        private void SendPeriodicHeartbeat()
        {
            try { Heartbeat(_lastNickname, _lastCellRef); }
            catch { }
        }

        /// <summary>TTL 만료 사용자/락 청소</summary>
        private void Sweep()
        {
            var now = DateTime.UtcNow;
            foreach (var kv in _presence)
            {
                if ((now - kv.Value).TotalSeconds > _ttlSec)
                {
                    _presence.TryRemove(kv.Key, out _);
                    ReleaseLocksBy(kv.Key);
                }
            }
        }

        // ===== Locks =====
        public bool TryAcquireCellLock(string sheetExAddr, string byNickname)
        {
            try
            {
                if (_cellLocks.TryAdd(sheetExAddr, byNickname))
                {
                    _backend?.AcquireLock(sheetExAddr, byNickname);
                    return true;
                }
                return false;
            }
            catch { return false; }
        }

        public void ReleaseLocksBy(string nickname)
        {
            try
            {
                foreach (var kv in _cellLocks)
                    if (kv.Value == nickname) _cellLocks.TryRemove(kv.Key, out _);

                _backend?.ReleaseLocksBy(nickname);
            }
            catch { }
        }

        public bool IsLocked(string sheetExAddr, out string by)
        {
            if (_cellLocks.TryGetValue(sheetExAddr, out by!)) return true;
            by = "";
            return false;
        }

        // ===== Backend (GraphQL.Client) =====
        private sealed class Backend : IDisposable
        {
            // 🔧 프로젝트 스키마에 맞게 필요시 이름/필드 교체
            private const string MUT_HEARTBEAT =
@"
mutation($nickname:String!, $cell:String){
  presenceHeartbeat(nickname:$nickname, cell:$cell) { ok }
}";
            private const string MUT_ACQUIRE =
@"
mutation($cell:String!, $by:String!){
  acquireLock(cell:$cell, by:$by) { ok }
}";
            private const string MUT_RELEASE_BY =
@"
mutation($by:String!){
  releaseLocksBy(by:$by) { ok }
}";

            private readonly GraphQLHttpClient _http;

            public Backend(Uri httpEndpoint, string apiKey)
            {
                _http = new GraphQLHttpClient(
                    new GraphQLHttpClientOptions { EndPoint = httpEndpoint },
                    new NewtonsoftJsonSerializer());
                if (!string.IsNullOrWhiteSpace(apiKey))
                    _http.HttpClient.DefaultRequestHeaders.Add("x-api-key", apiKey);
            }

            public void Dispose()
            {
                try { _http.Dispose(); } catch { }
            }

            public void PresenceHeartbeat(string nickname, string? cell)
            {
                try
                {
                    var req = new GraphQLRequest { Query = MUT_HEARTBEAT, Variables = new { nickname, cell } };
                    _http.SendMutationAsync<JObject>(req).GetAwaiter().GetResult();
                }
                catch { }
            }

            public void AcquireLock(string cell, string by)
            {
                try
                {
                    var req = new GraphQLRequest { Query = MUT_ACQUIRE, Variables = new { cell, by } };
                    _http.SendMutationAsync<JObject>(req).GetAwaiter().GetResult();
                }
                catch { }
            }

            public void ReleaseLocksBy(string by)
            {
                try
                {
                    var req = new GraphQLRequest { Query = MUT_RELEASE_BY, Variables = new { by } };
                    _http.SendMutationAsync<JObject>(req).GetAwaiter().GetResult();
                }
                catch { }
            }
        }
    }
}
