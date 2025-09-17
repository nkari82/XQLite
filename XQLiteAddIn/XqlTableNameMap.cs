using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace XQLite.AddIn
{
    public static class XqlTableNameMap
    {
        // 규칙: 소문자화 → 선행 접두 제거(Table_, tbl_, t_) → 비영숫자 → '_'
        public static string Normalize(string excelListObjectName, string? worksheetName = null)
        {
            var s = (excelListObjectName ?? string.Empty).Trim();
            if (string.IsNullOrEmpty(s)) s = worksheetName ?? string.Empty;
            s = s.Trim();
            s = s.StartsWith("Table_", StringComparison.OrdinalIgnoreCase) ? s.Substring(6) : s;
            s = s.StartsWith("tbl_", StringComparison.OrdinalIgnoreCase) ? s.Substring(4) : s;
            s = s.StartsWith("t_", StringComparison.OrdinalIgnoreCase) ? s.Substring(2) : s;
            var chars = s.Select(ch => char.IsLetterOrDigit(ch) ? char.ToLowerInvariant(ch) : '_').ToArray();
            var norm = new string(chars);
            // 연속 '_' 축소
            while (norm.Contains("__")) norm = norm.Replace("__", "_");
            return norm.Trim('_');
        }

        // 추가 매핑: %APPDATA%/XQLite/tablemap.json 지원  { "excelName":"server_table" }
        private static Dictionary<string, string>? _map;
        public static string Map(string excelListObjectName, string? worksheetName)
        {
            _map ??= LoadMap();
            if (_map is not null && _map.TryGetValue(excelListObjectName, out var v)) return v;
            return Normalize(excelListObjectName, worksheetName);
        }

        private static Dictionary<string, string>? LoadMap()
        {
            try
            {
                var dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "XQLite");
                var path = Path.Combine(dir, "tablemap.json");
                if (!File.Exists(path)) return new();
                var json = File.ReadAllText(path);
                var dict = XqlJson.Deserialize<Dictionary<string, string>>(json) ?? new();
                // 키를 Excel ListObject 이름 기준의 대/소문자 구분없이 취급
                return new Dictionary<string, string>(dict, StringComparer.OrdinalIgnoreCase);
            }
            catch { return new(); }
        }
    }
}