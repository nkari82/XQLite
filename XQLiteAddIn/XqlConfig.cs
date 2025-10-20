using ExcelDna.Integration;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.IO;
using System.Net;
using Excel = Microsoft.Office.Interop.Excel;


namespace XQLite.AddIn
{
    internal static class XqlConfig
    {
        [JsonProperty]
        public static string Endpoint { get; set; } = "http://localhost:4000/graphql";
        [JsonProperty]
        public static string ApiKey { get; set; } = "";
        [JsonProperty]
        public static string Nickname { get; set; } = Environment.UserName ?? "anonymous";
        [JsonProperty]
        public static string Project { get; set; } = "";


        [JsonProperty]
        public static int PullSec { get; set; } = 10;
        [JsonProperty]
        public static int DebounceMs { get; set; } = 2000;
        [JsonProperty]
        public static int HeartbeatSec { get; set; } = 3;
        [JsonProperty]
        public static int LockTtlSec { get; set; } = 10;

        [JsonProperty]
        public static bool AlwaysFullPullOnStartup { get; set; } = true;
        [JsonProperty]
        public static bool FullPullWhenSchemaChanged { get; set; } = true;
        [JsonProperty]
        public static string StateDirName { get; set; } = ".xql";

        private static string? _resolvedPath;


        public static void Load()
        {
            if (TryEnv("XQL_CONFIG"))
                return;

            var sidecar = TryWorkbookSidecar();
            if (sidecar is not null && TryFile(sidecar))
                return;

            var roaming = RoamingPath();
            if (TryFile(roaming))
                return;

            _resolvedPath = sidecar ?? roaming;
        }


        internal static void Save(bool preferSidecar = true)
        {
            var sidecar = TryWorkbookSidecar();
            var path = preferSidecar && sidecar is not null ? sidecar : RoamingPath();
            Save(path);
        }

        internal static void Save(string path)
        {
            var data = new
            {
                Endpoint,
                ApiKey,
                Nickname,
                Project,
                PullSec,
                DebounceMs,
                HeartbeatSec,
                LockTtlSec,
                AlwaysFullPullOnStartup,
                FullPullWhenSchemaChanged,
                StateDirName,
            };

            var json = JsonConvert.SerializeObject(data, Formatting.Indented);
            var dir = Path.GetDirectoryName(path);
            if (!string.IsNullOrEmpty(dir))
                Directory.CreateDirectory(dir!);
            File.WriteAllText(path, json);
            _resolvedPath = path;
        }

        private static bool TryEnv(string name)
        {
            var v = Environment.GetEnvironmentVariable(name);
            if (string.IsNullOrWhiteSpace(v)) return false;
            if (v.TrimStart().StartsWith("{"))
                return TryJson(v);
            if (File.Exists(v))
                return TryFile(v);
            return false;
        }

        internal static bool TryFile(string path)
        {
            try
            {
                return TryJson(File.ReadAllText(path)) && (_resolvedPath = path) == path;
            }
            catch { return false; }
        }

        internal static bool TryJson(string json)
        {
            try
            {
                var x = JObject.Parse(json);
                if (x == null)
                    return false;

                Endpoint = (string?)x["Endpoint"] ?? Endpoint;
                ApiKey = (string?)x["ApiKey"] ?? ApiKey;
                Nickname = (string?)x["Nickname"] ?? Nickname;
                Project = (string?)x["Project"] ?? Project;
                PullSec = (int?)x["PullSec"] ?? PullSec;
                DebounceMs = (int?)x["DebounceMs"] ?? DebounceMs;
                HeartbeatSec = (int?)x["HeartbeatSec"] ?? HeartbeatSec;
                LockTtlSec = (int?)x["LockTtlSec"] ?? LockTtlSec;
                AlwaysFullPullOnStartup = (bool?)x["AlwaysFullPullOnStartup"] ?? AlwaysFullPullOnStartup;
                FullPullWhenSchemaChanged = (bool?)x["FullPullWhenSchemaChanged"] ?? FullPullWhenSchemaChanged;
                StateDirName = (string?)x["StateDirName"] ?? StateDirName;

                return true;
            }
            catch
            {
                return false;
            }
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