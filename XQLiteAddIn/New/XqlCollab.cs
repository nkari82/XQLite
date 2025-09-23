// XqlCollab.cs
using System;
using System.Collections.Concurrent;
using System.Threading;
using System.Threading.Tasks;
using GraphQL;
using GraphQL.Client.Http;
using GraphQL.Client.Serializer.Newtonsoft;
using Newtonsoft.Json.Linq;

namespace XQLite.AddIn
{
    /// <summary>
    /// 실시간 협업(프레즌스 + 시트/셀/컬럼 락)
    /// - Heartbeat: 2.5초 주기(기본), TTL: 10초(기본)
    /// - 로컬 캐시에 즉시 반영 + 서버와 동기화
    /// - 락 키 규격:
    ///   * 셀   : "cell:{sheet}!{address}"
    ///   * 컬럼 : "column:{table}.{column}"
    /// </summary>
    internal sealed class XqlCollab : IDisposable
    {
        private readonly ConcurrentDictionary<string, DateTime> _presence = new(StringComparer.Ordinal);
        private readonly ConcurrentDictionary<string, string> _locks = new(StringComparer.Ordinal); // key -> nickname

        private readonly Timer _ttlSweep;
        private readonly Timer _hbTimer;

        private readonly IXqlBackend? _backend; // 서버 연동(선택)

        private readonly int _ttlSec;
        private readonly int _hbMs;

        private volatile string _lastNickname = "anonymous";
        private volatile string _lastCellRef = ""; // "Sheet!A1"

        internal static XqlCollab? Instance = null;

        public XqlCollab(IXqlBackend backend, int ttlSeconds = 10, int heartbeatMs = 2500)
        {
            Instance = null;

            _backend = backend ?? throw new ArgumentNullException(nameof(backend));
            _ttlSec = Math.Max(5, ttlSeconds);
            _hbMs = Math.Max(1000, heartbeatMs);

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

        /// <summary>선택 변경 시 호출: 현재 선택 셀을 알려주면 디바운스된 하트비트로 반영됩니다.</summary>
        public void NotifySelection(string nickname, string sheetExAddr)
        {
            _lastNickname = string.IsNullOrWhiteSpace(nickname) ? "anonymous" : nickname;
            _lastCellRef = sheetExAddr ?? "";
            _hbTimer.Change(_hbMs, Timeout.Infinite);
        }

        /// <summary>즉시 하트비트(버튼/메뉴에서 수동 호출 가능)</summary>
        public void Heartbeat(string nickname, string? cellRef = null)
        {
            try
            {
                nickname = string.IsNullOrWhiteSpace(nickname) ? "anonymous" : nickname;
                var cref = cellRef ?? _lastCellRef;

                _presence[nickname] = DateTime.UtcNow;
                if (!string.IsNullOrEmpty(cref))
                {
                    var key = BuildCellKeyFromDisplayRef(cref); // "cell:Sheet!A1"
                    _locks.TryAdd(key, nickname);
                }

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

        // ===== Locks: 공통 유틸 =====

        private static string BuildCellKey(string sheet, string address) => $"cell:{sheet}!{address}";
        private static string BuildCellKeyFromDisplayRef(string sheetExAddr) => $"cell:{sheetExAddr}";
        private static string BuildColumnKey(string table, string column) => $"column:{table}.{column}";

        /// <summary>소유자 닉네임 기준으로 모든 락 해제</summary>
        public void ReleaseLocksBy(string nickname)
        {
            try
            {
                foreach (var kv in _locks)
                    if (kv.Value == nickname) _locks.TryRemove(kv.Key, out _);

                _backend?.ReleaseLocksBy(nickname);
            }
            catch { }
        }

        public bool IsLockedCell(string sheet, string address, out string by)
        {
            var key = BuildCellKey(sheet, address);
            if (_locks.TryGetValue(key, out by!)) return true;
            by = ""; return false;
        }

        public bool IsLockedColumn(string table, string column, out string by)
        {
            var key = BuildColumnKey(table, column);
            if (_locks.TryGetValue(key, out by!)) return true;
            by = ""; return false;
        }

        /// <summary>셀 락 시도(동기). 성공 시 로컬/서버에 반영.</summary>
        public bool TryAcquireCellLock(string sheet, string address, string byNickname)
        {
            try
            {
                var key = BuildCellKey(sheet, address);
                if (_locks.TryAdd(key, byNickname))
                {
                    // 서버에는 통합 acquireLock(key, by)로 전달
                    _backend?.AcquireLock(key, byNickname);
                    return true;
                }
                return false;
            }
            catch { return false; }
        }

        /// <summary>컬럼 락 시도(동기). 성공 시 로컬/서버에 반영.</summary>
        public bool TryAcquireColumnLock(string table, string column, string byNickname)
        {
            try
            {
                var key = BuildColumnKey(table, column);
                if (_locks.TryAdd(key, byNickname))
                {
                    // 서버에는 통합 acquireLock(key, by)로 전달
                    _backend?.AcquireLock(key, byNickname);
                    return true;
                }
                return false;
            }
            catch { return false; }
        }
        public Task<bool> AcquireColumnAsync(string table, string column, int ttlSec = 10)
            => Task.Run(() => TryAcquireColumnLock(table, column, _lastNickname));

        /// <summary>
        /// 기존: XqlLockService.AcquireCellAsync(sheet, address)
        /// - 닉네임은 내부 Presence 상태(_lastNickname) 사용
        /// </summary>
        public Task<bool> AcquireCellAsync(string sheet, string address, int ttlSec = 10)
            => Task.Run(() => TryAcquireCellLock(sheet, address, _lastNickname));
    }
}
