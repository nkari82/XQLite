using System;
using System.IO;
using System.Windows.Forms;
using ExcelDna.Integration;
using Excel = Microsoft.Office.Interop.Excel;

namespace XQLite.AddIn
{
    internal sealed class XqlAddIn : IExcelAddIn
    {
        // ====== Config ======
        internal static XqlConfig? Cfg { get; set; }
        internal static IXqlBackend? Backend { get; }
        internal static XqlMetaRegistry? MetaRegistry { get; }
        internal static XqlSync? Sync { get; }
        internal static XqlCollab? Collab { get; }
        internal static XqlBackup? Backup { get; }
        internal static XqlExcelInterop? ExcelInterop { get; }

        private static readonly string AppDir =
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "XQLite");
        private static readonly string CfgPath = Path.Combine(AppDir, "config.json");

        // ====== Singletons (런타임 구성요소) ======
        private static IXqlBackend? _backend;
        private static XqlMetaRegistry? _meta;
        private static XqlSync? _sync;
        private static XqlCollab? _collab;
        private static XqlBackup? _backup;
        private static XqlExcelInterop? _interop;

        public void AutoOpen()
        {
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
        }

        public void AutoClose()
        {
            try { StopRuntime(); }
            catch { /* ignore */ }
        }

        // ====== Runtime lifecycle ======
        internal static void RestartRuntime(XqlConfig cfg)
        {
            StopRuntime();
            StartRuntime(cfg);
        }

        internal static void StartRuntime(XqlConfig cfg)
        {
            try
            {
                // 1) 백엔드 & 메타
                _backend = new XqlGqlBackend(cfg.Endpoint, cfg.ApiKey); // GraphQL 클라이언트(HTTP/WS) 공용 인스턴스
                _meta = new XqlMetaRegistry();

                // 2) 동기화/협업/백업
                //    - XqlSync: push(업서트) ms, pull(증분) ms
                _sync = new XqlSync(_backend, _meta,
                    pushIntervalMs: Math.Max(250, cfg.DebounceMs),
                    pullIntervalMs: Math.Max(1000, cfg.PullSec * 1000)); // Start/Stop 지원
                _sync.Start(); // 구독 시작 포함 :contentReference[oaicite:5]{index=5}

                //    - XqlCollab: TTL/Heartbeat 간격 (초→ms)
                _collab = new XqlCollab(_backend,
                    ttlSeconds: Math.Max(5, cfg.LockTtlSec),
                    heartbeatMs: Math.Max(1000, cfg.HeartbeatSec * 1000)); // 내부 타이머 운용 :contentReference[oaicite:6]{index=6}

                //    - XqlBackup: 진단/복구/풀덤프
                _backup = new XqlBackup(_backend, _meta, cfg.Endpoint, cfg.ApiKey); // 현재 시그니처 기준 :contentReference[oaicite:7]{index=7}

                // 3) Excel 이벤트 훅
                var app = (Excel.Application)ExcelDnaUtil.Application;
                _interop = new XqlExcelInterop(app, _sync, _collab, _meta, _backup);
                _interop.Start(); // Excel SheetChange/SelectionChange/Workbook 이벤트 연결 :contentReference[oaicite:8]{index=8}
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to start runtime: " + ex.Message, "XQLite", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        internal static void StopRuntime()
        {
            try
            {
                // Excel 이벤트 해제
                try { _interop?.Stop(); } catch { }
                try { _interop?.Dispose(); } catch { }
                _interop = null;

                // 동기화/협업 정리
                try { _sync?.Stop(); } catch { }
                try { _sync?.Dispose(); } catch { }
                _sync = null;

                try { _collab?.Dispose(); } catch { }
                _collab = null;

                // 백업 객체 정리
                try { _backup?.Dispose(); } catch { }
                _backup = null;

                // 백엔드 정리(마지막에)
                try { _backend?.Dispose(); } catch { }
                _backend = null;

                // 메타는 관리 객체이므로 GC에 맡김
                _meta = null;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to stop runtime: " + ex.Message, "XQLite", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        // ====== Config load/save (Env → File → Defaults) ======
        private static XqlConfig LoadConfigWithOverrides()
        {
            var cfg = LoadConfigFromFile() ?? new XqlConfig();

            // 1) 환경변수 우선 적용
            string? ep = Environment.GetEnvironmentVariable("XQLITE_ENDPOINT");
            if (!string.IsNullOrWhiteSpace(ep)) cfg.Endpoint = ep.Trim();

            string? k = Environment.GetEnvironmentVariable("XQLITE_APIKEY");
            if (!string.IsNullOrWhiteSpace(k)) cfg.ApiKey = k.Trim();

            string? nick = Environment.GetEnvironmentVariable("XQLITE_NICKNAME");
            if (!string.IsNullOrWhiteSpace(nick)) cfg.Nickname = nick.Trim();

            string? proj = Environment.GetEnvironmentVariable("XQLITE_PROJECT");
            if (!string.IsNullOrWhiteSpace(proj)) cfg.Project = proj.Trim();

            // 2) 기본값 보정
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
}
