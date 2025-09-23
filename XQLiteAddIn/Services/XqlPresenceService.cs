using System;
using System.Threading.Tasks;
using System.Timers;

#if false
namespace XQLite.AddIn
{
    internal static class XqlPresenceService
    {
        private static Timer? _hb;
        private static XqlConfig? _cfg;


        internal static void Start(XqlConfig cfg)
        {
            _cfg = cfg;
            _hb = new Timer(Math.Max(1000, cfg.HeartbeatSec * 1000)) { AutoReset = true };
            _hb.Elapsed += async (_, __) => await TickAsync();
            _hb.Start();
        }


        internal static void Stop()
        {
            if (_hb is not null) 
            { 
                _hb.Stop(); 
                _hb.Dispose(); 
                _hb = null; 
            }
            _cfg = null;
        }


        private static async Task TickAsync()
        {
            if (_cfg is null) 
                return;

            const string q = "mutation($nick:String!){ presenceHeartbeat(nickname:$nick){ ok ttl } }";
            try 
            { 
                await XqlGraphQLClient.MutateAsync<dynamic>(q, new { nick = _cfg.Nickname }); 
            }
            catch(Exception)
            {
                /* 네트워크 오류 무시(다음 틱에 재시도) */
            }
        }
    }
}
#endif