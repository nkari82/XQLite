using ExcelDna.Integration;
using System;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;


namespace XQLite.AddIn
{
    internal sealed class XqlConfig
    {
        public string Endpoint { get; set; } = "http://localhost:4000/graphql";
        public string ApiKey { get; set; } = "";
        public string Nickname { get; set; } = Environment.UserName;
        public string Project { get; set; } = "";


        public int PullSec { get; set; } = 10;
        public int DebounceMs { get; set; } = 2000;
        public int HeartbeatSec { get; set; } = 3;
        public int LockTtlSec { get; set; } = 10;


        private string? _resolvedPath;


        public static XqlConfig Load()
        {
            var cfg = new XqlConfig();
            if (TryEnv("XQL_CONFIG", ref cfg)) return cfg;


            var sidecar = TryWorkbookSidecar();
            if (sidecar is not null && TryFile(sidecar, ref cfg)) return cfg;


            var roaming = RoamingPath();
            if (TryFile(roaming, ref cfg)) return cfg;

            cfg._resolvedPath = sidecar ?? roaming;
            return cfg;
        }


        public void Save(bool preferSidecar = true)
        {
            var sidecar = TryWorkbookSidecar();
            var path = preferSidecar && sidecar is not null ? sidecar : RoamingPath();
            var json = XqlJson.Serialize(this, true);
            var dir = Path.GetDirectoryName(path);
            if(!string.IsNullOrEmpty(dir))
                Directory.CreateDirectory(dir!);
            File.WriteAllText(path, json);
            _resolvedPath = path;
        }


        private static bool TryEnv(string name, ref XqlConfig cfg)
        {
            var v = Environment.GetEnvironmentVariable(name);
            if (string.IsNullOrWhiteSpace(v)) return false;
            if (v.TrimStart().StartsWith("{"))
                return TryJson(v, ref cfg);
            if (File.Exists(v))
                return TryFile(v, ref cfg);
            return false;
        }

        private static bool TryFile(string path, ref XqlConfig cfg)
        {
            try { return TryJson(File.ReadAllText(path), ref cfg) && (cfg._resolvedPath = path) == path; }
            catch { return false; }
        }

        private static bool TryJson(string json, ref XqlConfig cfg)
        {
            try { var x = XqlJson.Deserialize<XqlConfig>(json); if (x is null) return false; Copy(x, cfg); return true; }
            catch { return false; }
        }

        private static void Copy(XqlConfig s, XqlConfig d)
        {
            d.Endpoint = s.Endpoint; d.ApiKey = s.ApiKey; d.Nickname = s.Nickname; d.Project = s.Project;
            d.PullSec = s.PullSec; d.DebounceMs = s.DebounceMs; d.HeartbeatSec = s.HeartbeatSec; d.LockTtlSec = s.LockTtlSec;
        }

        private static string RoamingPath()
        {
            var dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "XQLite");
            return Path.Combine(dir, "config.json");
        }

        private static string? TryWorkbookSidecar()
        {
            try
            {
                var app = (Excel.Application)ExcelDnaUtil.Application;
                var wb = (Excel.Workbook)app.ActiveWorkbook;
                var full = wb?.FullName;
                if (string.IsNullOrWhiteSpace(full)) return null;
                return Path.ChangeExtension(full, ".xql.json");
            }
            catch { return null; }
        }
    }
}