using System.Text.RegularExpressions;

namespace XQLite.AddIn
{
    public static class XqlPathParser
    {
        // 예상 예: "rows[3].price" or "rows[10].meta.name"
        private static readonly Regex RxIndex = new(@"rows\[(\d+)\](?:\.([A-Za-z0-9_\.]+))?", RegexOptions.Compiled);

        public static (int? rowIndex1, string? columnPath) Parse(string? path)
        {
            if (string.IsNullOrWhiteSpace(path)) return (null, null);
            var m = RxIndex.Match(path);
            if (!m.Success) return (null, null);
            int idx = int.Parse(m.Groups[1].Value);
            string? col = m.Groups.Count > 2 ? m.Groups[2].Value : null;
            return (idx + 1, string.IsNullOrWhiteSpace(col) ? null : col); // 1-based로 표시
        }
    }
}