// XqlDiagExport.cs
using ExcelDna.Integration;
using Microsoft.Win32;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace XQLite.AddIn
{
    /// <summary>
    /// 진단용 ZIP 번들을 생성한다.
    /// - info.json : OS/Excel/Add-in/프로세스 정보
    /// - config.json : 애드인 설정(있으면)
    /// - logs/*.log : 파일 로그(있으면)
    /// - outbox.ndjson : 업서트 실패 아웃박스(있으면)
    /// - graph/*.json : 서버 health, presence, (선택)rows 일부
    /// - xl/selection.txt : 현재 선택 정보
    /// </summary>
    internal static class XqlDiagExport
    {
        // 기본 위치: %APPDATA%\XQLite\diag\XQLite_Diag_yyyyMMdd_HHmmss.zip
        internal static string DefaultDiagZipPath()
        {
            var root = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "XQLite", "diag");
            Directory.CreateDirectory(root);
            var ts = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            return Path.Combine(root, $"XQLite_Diag_{ts}.zip");
        }

        /// <summary>
        /// 지정 경로로 진단 ZIP 생성
        /// </summary>
        internal static async Task ExportAsync(string zipPath, CancellationToken ct = default)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(zipPath) ?? ".");
            if (File.Exists(zipPath))
                File.Delete(zipPath);

            using var fs = new FileStream(zipPath, FileMode.CreateNew, FileAccess.ReadWrite, FileShare.None);
            using var zip = new ZipArchive(fs, ZipArchiveMode.Create, leaveOpen: false, entryNameEncoding: Encoding.UTF8);
            // 1) info.json
            var info = CollectInfo();
            AddJson(zip, "info.json", info);

            // 2) config.json (있으면)
            TryAddConfig(zip);

            // 3) logs/*
            TryAddLogs(zip);

            // 4) outbox.ndjson (있으면)
            TryAddOutbox(zip);

            // 5) Excel 상태
            TryAddExcelSelection(zip);

            // 6) GraphQL 진단 호출들
            await TryAddGraphQLDiagAsync(zip, ct).ConfigureAwait(false);
        }

        // --------------------------
        // 수집기들
        // --------------------------

        private static JObject CollectInfo()
        {
            string excelVer = "-", officeBit = "-";
            try
            {
                var app = (Excel.Application)ExcelDnaUtil.Application;
                excelVer = app.Version;
                officeBit = Environment.Is64BitProcess ? "x64" : "x86";
            }
            catch { /* Excel 미첨부 시 무시 */ }

            var asm = Assembly.GetExecutingAssembly();
            var asmName = asm.GetName();

            var info = new JObject
            {
                ["time"] = DateTime.Now.ToString("o"),
                ["os"] = new JObject
                {
                    ["machineName"] = Environment.MachineName,
                    ["userName"] = Environment.UserName,
                    ["osVersion"] = Environment.OSVersion.VersionString,
                    ["is64BitOS"] = Environment.Is64BitOperatingSystem,
                },
                ["process"] = new JObject
                {
                    ["is64BitProcess"] = Environment.Is64BitProcess,
                    ["clr"] = Environment.Version.ToString(),
                    ["cwd"] = Environment.CurrentDirectory,
                    ["cmdLine"] = Environment.CommandLine
                },
                ["excel"] = new JObject
                {
                    ["version"] = excelVer,
                    ["bitness"] = officeBit
                },
                ["addin"] = new JObject
                {
                    ["assembly"] = asmName.Name,
                    ["version"] = asmName.Version?.ToString(),
                    ["location"] = SafeGet(() => asm.Location)
                }
            };
            return info;
        }

        private static void TryAddConfig(ZipArchive zip)
        {
            // 프로젝트 쪽에서 쓰는 설정 저장소를 모르는 경우 대비: 흔한 경로들 점검
            // 1) %APPDATA%\XQLite\config.json
            var appData = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "XQLite");
            var candidate = Path.Combine(appData, "config.json");
            if (File.Exists(candidate))
            {
                AddFile(zip, "config/config.json", candidate);
                return;
            }

            // 2) XqlConfig.Current 같은 정적 제공자가 있다면 직렬화 시도
            try
            {
                var tConfig = Type.GetType("XQLite.AddIn.XqlConfig, XQLite.AddIn", throwOnError: false);
                if (tConfig != null)
                {
                    var prop = tConfig.GetProperty("Current", BindingFlags.Public | BindingFlags.Static);
                    var cur = prop?.GetValue(null, null);
                    if (cur != null)
                    {
                        AddJson(zip, "config/current.json", JObject.FromObject(cur));
                    }
                }
            }
            catch { /* optional */ }
        }

        private static void TryAddLogs(ZipArchive zip)
        {
            // XqlFileLogger가 LogDirectory를 제공하도록 이전에 구현해 두었다면 사용
#pragma warning disable CS8600 // null 리터럴 또는 가능한 null 값을 null을 허용하지 않는 형식으로 변환하는 중입니다.
            var dir = SafeGet(static () => (string)typeof(XqlFileLogger).GetProperty("LogDirectory", BindingFlags.Public | BindingFlags.Static)?.GetValue(null, null));
#pragma warning restore CS8600 // null 리터럴 또는 가능한 null 값을 null을 허용하지 않는 형식으로 변환하는 중입니다.
            if (string.IsNullOrEmpty(dir))
                dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "XQLite", "logs");

            if (!Directory.Exists(dir)) return;

            foreach (var file in Directory.GetFiles(dir, "*.log"))
                AddFile(zip, "logs/" + Path.GetFileName(file), file);
        }

        private static void TryAddOutbox(ZipArchive zip)
        {
            var path = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "XQLite", "outbox.ndjson");
            if (File.Exists(path))
                AddFile(zip, "outbox/outbox.ndjson", path);
        }

        private static void TryAddExcelSelection(ZipArchive zip)
        {
            try
            {
                var app = (Excel.Application)ExcelDnaUtil.Application;
                var ws = (Excel.Worksheet)app.ActiveSheet;
                var sel = (Excel.Range)app.Selection;
                string addr = SafeGet(() => sel?.Address[false, false]) ?? "-";
                string text = $"Sheet: {ws?.Name ?? "-"}{Environment.NewLine}Selection: {addr}";

                AddText(zip, "xl/selection.txt", text);
            }
            catch
            {
                // Excel 없이 실행되는 경우 무시
            }
        }

        private static async Task TryAddGraphQLDiagAsync(ZipArchive zip, CancellationToken ct)
        {
            // 가벼운 질의만
            // 1) health
            try
            {
                var qr = await XqlGraphQLClient.QueryAsync<JObject>("query{ health }", null, ct).ConfigureAwait(false);
                AddJson(zip, "graph/health.json", JObject.FromObject(qr));
            }
            catch (Exception ex)
            {
                AddText(zip, "graph/health.error.txt", ex.ToString());
            }

            // 2) presence
            try
            {
                var qr = await XqlGraphQLClient.QueryAsync<JObject>("query{ presence{ nickname sheet cell updated_at } }", null, ct).ConfigureAwait(false);
                AddJson(zip, "graph/presence.json", JObject.FromObject(qr));
            }
            catch (Exception ex)
            {
                AddText(zip, "graph/presence.error.txt", ex.ToString());
            }

            // 3) rows — 과도한 데이터 방지: 각 블록에서 최대 30행 정도만 샘플링 (서버가 rows(since)만 제공한다고 가정)
            try
            {
                var qr = await XqlGraphQLClient.QueryAsync<JObject>("query($since:Long){ rows(since_version:$since){ table rows max_row_version } }", new { since = 0L }, ct).ConfigureAwait(false);
                var root = JObject.FromObject(qr);
                var rows = root["Data"]?["rows"] as JArray ?? root["data"]?["rows"] as JArray;
                if (rows != null)
                {
                    var capped = new JArray();
                    foreach (var blk in rows.OfType<JObject>())
                    {
                        var table = blk["table"]?.ToString() ?? "unknown";
                        var arr = blk["rows"] as JArray ?? new JArray();
                        var sample = new JArray(arr.Take(30)); // cap 30
                        var obj = new JObject
                        {
                            ["table"] = table,
                            ["rows"] = sample,
                            ["max_row_version"] = blk["max_row_version"] ?? 0
                        };
                        capped.Add(obj);
                    }
                    AddJson(zip, "graph/rows_sample.json", capped);
                }
                else
                {
                    AddJson(zip, "graph/rows_raw.json", root);
                }
            }
            catch (Exception ex)
            {
                AddText(zip, "graph/rows.error.txt", ex.ToString());
            }
        }

        // --------------------------
        // ZIP 헬퍼들
        // --------------------------
        private static void AddText(ZipArchive zip, string entry, string text)
        {
            var e = zip.CreateEntry(entry, CompressionLevel.Optimal);
            using var s = e.Open();
            using var sw = new StreamWriter(s, Encoding.UTF8);
            sw.Write(text ?? string.Empty);
        }

        private static void AddJson(ZipArchive zip, string entry, object obj)
        {
            var json = JsonConvert.SerializeObject(obj, Formatting.Indented);
            AddText(zip, entry, json);
        }

        private static void AddFile(ZipArchive zip, string entry, string filePath)
        {
            var e = zip.CreateEntry(entry, CompressionLevel.Optimal);
            using var src = File.OpenRead(filePath);
            using var dst = e.Open();
            src.CopyTo(dst);
        }

        private static T SafeGet<T>(Func<T> f)
        {
#pragma warning disable CS8603 // 가능한 null 참조 반환입니다.
            try { return f(); } catch { return default(T); }
#pragma warning restore CS8603 // 가능한 null 참조 반환입니다.
        }
    }
}
