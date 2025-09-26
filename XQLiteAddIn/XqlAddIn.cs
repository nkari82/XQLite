using System;
using System.IO;
using System.Windows.Forms;
using ExcelDna.Integration;
using Excel = Microsoft.Office.Interop.Excel;

namespace XQLite.AddIn
{
    internal sealed class XqlAddIn : IExcelAddIn
    {
        internal static IXqlBackend? Backend => _backend;
        internal static XqlSheet? Sheet => _sheet;
        internal static XqlSync? Sync => _sync;
        internal static XqlCollab? Collab => _collab;
        internal static XqlBackup? Backup => _backup;
        internal static XqlExcelInterop? ExcelInterop => _interop;

        private static readonly string AppDir =
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "XQLite");
        private static readonly string CfgPath = Path.Combine(AppDir, "config.json");

        // ====== Singletons (런타임 구성요소) ======
        private static IXqlBackend? _backend;
        private static XqlSheet? _sheet;
        private static XqlSync? _sync;
        private static XqlCollab? _collab;
        private static XqlBackup? _backup;
        private static XqlExcelInterop? _interop;

        public void AutoOpen()
        {
            try
            {
                Directory.CreateDirectory(AppDir);
                LoadConfigWithOverrides();
                StartRuntime();
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
        internal static void RestartRuntime()
        {
            StopRuntime();
            StartRuntime();
        }

        internal static void StartRuntime()
        {
            try
            {
                // 1) 백엔드 & 메타
                _backend = new XqlGqlBackend(XqlConfig.Endpoint, XqlConfig.ApiKey); // GraphQL 클라이언트(HTTP/WS) 공용 인스턴스
                _sheet = new XqlSheet();

                // 2) 동기화/협업/백업
                //    - XqlSync: push(업서트) ms, pull(증분) ms
                _sync = new XqlSync(_backend, _sheet,
                    pushIntervalMs: Math.Max(250, XqlConfig.DebounceMs),
                    pullIntervalMs: Math.Max(1000, XqlConfig.PullSec * 1000)); // Start/Stop 지원
                _sync.Start(); // 구독 시작 포함

                //    - XqlCollab: TTL/Heartbeat 간격 (초→ms)
                _collab = new XqlCollab(_backend, XqlConfig.Nickname, heartbeatSec: XqlConfig.HeartbeatSec); // 내부 타이머 운용

                //    - XqlBackup: 진단/복구/풀덤프
                _backup = new XqlBackup(_backend, _sheet, XqlConfig.Endpoint, XqlConfig.ApiKey); // 현재 시그니처 기준

                // 3) Excel 이벤트 훅
                var app = (Excel.Application)ExcelDnaUtil.Application;
                _interop = new XqlExcelInterop(app, _sync, _collab, _sheet, _backup);
                _interop.Start(); // Excel SheetChange/SelectionChange/Workbook 이벤트 연결
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
                _sheet = null;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to stop runtime: " + ex.Message, "XQLite", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        // ====== Config load/save (Env → File → Defaults) ======
        private static void LoadConfigWithOverrides()
        {
            LoadConfigFromFile();

            // 1) 환경변수 우선 적용
            string? ep = Environment.GetEnvironmentVariable("XQLITE_ENDPOINT");
            if (!string.IsNullOrWhiteSpace(ep)) XqlConfig.Endpoint = ep.Trim();

            string? k = Environment.GetEnvironmentVariable("XQLITE_APIKEY");
            if (!string.IsNullOrWhiteSpace(k)) XqlConfig.ApiKey = k.Trim();

            string? nick = Environment.GetEnvironmentVariable("XQLITE_NICKNAME");
            if (!string.IsNullOrWhiteSpace(nick)) XqlConfig.Nickname = nick.Trim();

            string? proj = Environment.GetEnvironmentVariable("XQLITE_PROJECT");
            if (!string.IsNullOrWhiteSpace(proj)) XqlConfig.Project = proj.Trim();

            // 2) 기본값 보정
            XqlConfig.PullSec = XqlConfig.PullSec <= 0 ? 10 : XqlConfig.PullSec;
            XqlConfig.DebounceMs = XqlConfig.DebounceMs <= 0 ? 2000 : XqlConfig.DebounceMs;
            XqlConfig.HeartbeatSec = XqlConfig.HeartbeatSec <= 0 ? 3 : XqlConfig.HeartbeatSec;
            XqlConfig.LockTtlSec = XqlConfig.LockTtlSec <= 0 ? 10 : XqlConfig.LockTtlSec;
        }

        private static void LoadConfigFromFile()
        {
            try
            {
                if (!File.Exists(CfgPath)) return;
                var json = File.ReadAllText(CfgPath);
                XqlConfig.TryJson(json);
            }
            catch { }
        }

        internal static void SaveConfigToFile()
        {
            try
            {
                Directory.CreateDirectory(AppDir);
                XqlConfig.Save(CfgPath);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to save config: " + ex.Message, "XQLite", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
    }
}
