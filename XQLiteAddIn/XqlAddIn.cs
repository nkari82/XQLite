// XqlAddIn.cs (no partial)
// - Excel-DNA IExcelAddIn 엔트리포인트를 단일 파일로 통합
// - AutoOpen/AutoClose에서 런타임 시작/정지
// - 리본/단축키에서 부르는 공개 명령 메서드(Commit/Recover/Inspector/Export/Presence/Schema/Diag)
// - 정책: 시트엔 절대 쓰지 않음(표시/색칠/코멘트 X)

using System;
using System.IO;
using System.Windows.Forms;
using ExcelDna.Integration;

namespace XQLite.AddIn
{
    public sealed class XqlAddIn : IExcelAddIn
    {
        // ====== Config ======
        public static XqlConfig? Cfg { get; internal set; }
        private static readonly string AppDir =
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "XQLite");
        private static readonly string CfgPath = Path.Combine(AppDir, "config.json");

        public void AutoOpen()
        {
#if true
            try
            {
                Directory.CreateDirectory(AppDir);
                Cfg = LoadConfigWithOverrides();
                StartRuntime(Cfg!);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"XQLite failed to start:\r\n{ex}", "XQLite", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
#endif
        }

        public void AutoClose()
        {
#if true
            try
            {
                StopRuntime();
            }
            catch { /* ignore */ }
#endif
        }

        // ====== Runtime lifecycle ======
        internal static void RestartRuntime(XqlConfig cfg)
        {
            StopRuntime();
            StartRuntime(cfg);
        }

        internal static void StartRuntime(XqlConfig cfg)
        {
            // 1) 기본 서비스 초기화
            XqlGraphQLClient.Init(cfg);
            XqlPresenceService.Start(cfg);

            // 2) 락/동기화/구독 (필요 정책대로 택1/병행)
            XqlLockService.Start();

            XqlUpsert.Init(cfg.DebounceMs);
            XqlSubscriptionService.Start(cfg, startSince: 0);

            // 3) 시트 이벤트 후킹
            XqlSheetEvents.Hook();

            // 4) 파일 로거(부트스트랩에서 이미 켰다면 생략 가능)
            XqlFileLogger.Start();
        }

        internal static void StopRuntime()
        {
            XqlSheetEvents.Unhook();
            try { XqlSubscriptionService.Stop(); } catch { }
            try { XqlPresenceService.Stop(); } catch { }
            try { XqlLockService.Stop(); } catch { }

            // 3) 파일 로거 종료(부트스트랩에서 관리 중이면 생략)
            try { XqlFileLogger.Stop(); } catch { }
        }

        // ====== Config load/save (Env → File → Defaults) ======
        private static XqlConfig LoadConfigWithOverrides()
        {
            var cfg = LoadConfigFromFile() ?? new XqlConfig();

            // 1) 환경변수 우선 적용
            //    (이름이 다르면 적합하게 교체)
            string? ep = Environment.GetEnvironmentVariable("XQLITE_ENDPOINT");
            if (!string.IsNullOrWhiteSpace(ep)) cfg.Endpoint = ep.Trim();

            string? k = Environment.GetEnvironmentVariable("XQLITE_APIKEY");
            if (!string.IsNullOrWhiteSpace(k)) cfg.ApiKey = k.Trim();

            string? nick = Environment.GetEnvironmentVariable("XQLITE_NICKNAME");
            if (!string.IsNullOrWhiteSpace(nick)) cfg.Nickname = nick.Trim();

            string? proj = Environment.GetEnvironmentVariable("XQLITE_PROJECT");
            if (!string.IsNullOrWhiteSpace(proj)) cfg.Project = proj.Trim();

            // 2) 기본값 보정 (없으면 안전한 값)
            cfg.PullSec = cfg.PullSec <= 0 ? 10 : cfg.PullSec;
            cfg.DebounceMs = cfg.DebounceMs <= 0 ? 2000 : cfg.DebounceMs;
            cfg.HeartbeatSec = cfg.HeartbeatSec <= 0 ? 3 : cfg.HeartbeatSec;
            cfg.LockTtlSec = cfg.LockTtlSec <= 0 ? 10 : cfg.LockTtlSec;

            return cfg;
        }

        private static XqlConfig? LoadConfigFromFile()
        {
            try
            {
                if (!File.Exists(CfgPath)) return null;
                var json = File.ReadAllText(CfgPath);
                return XqlJson.Deserialize<XqlConfig>(json);
            }
            catch { return null; }
        }

        internal static void SaveConfigToFile(XqlConfig cfg)
        {
            try
            {
                Directory.CreateDirectory(AppDir);
                var json = XqlJson.Serialize(cfg, true);
                File.WriteAllText(CfgPath, json);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to save config: " + ex.Message, "XQLite", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
    }

    internal static class XqlDiagExport
    {
        public static string ExportZip()
        {
            // TODO: STEP14에서 구현한 실제 Zip 수집 로직으로 교체
            var path = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                                    "XQLite", "diag", "xql_diag_dummy.zip");
            Directory.CreateDirectory(Path.GetDirectoryName(path)!);
            File.WriteAllBytes(path, Array.Empty<byte>());
            return path;
        }
    }
}
