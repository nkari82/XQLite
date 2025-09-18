using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;

namespace XQLite.AddIn
{
    internal static class XqlRowCache
    {
        // 캐시 키: table + '|' + keyValue  (keyValue가 없으면 행주소 키)
        private static readonly Dictionary<string, string> _hashByKey = new(StringComparer.Ordinal);

        // 딕셔너리 → 안정적 해시 문자열(정렬된 키, null/빈문자 규칙 포함)
        internal static string ComputeHash(IReadOnlyDictionary<string, object?> row)
        {
            var sb = new StringBuilder();
            foreach (var kv in row.OrderBy(k => k.Key, StringComparer.Ordinal))
            {
                sb.Append(kv.Key);
                sb.Append('=');
                sb.Append(ValueToStableString(kv.Value));
                sb.Append(';');
            }
            using var sha = SHA256.Create();
            var bytes = Encoding.UTF8.GetBytes(sb.ToString());
            var hash = sha.ComputeHash(bytes);
            return Compat.ToHexString(hash);
        }

        internal static bool IsChanged(string table, string rowKey, string hash)
        {
            var k = table + "|" + rowKey;
            if (!_hashByKey.TryGetValue(k, out var prev)) { _hashByKey[k] = hash; return true; }
            if (!string.Equals(prev, hash, StringComparison.Ordinal)) { _hashByKey[k] = hash; return true; }
            return false;
        }

        private static string ValueToStableString(object? v)
        {
            if (v is null) return "<null>";
            switch (v)
            {
                case double d: return d.ToString("R", System.Globalization.CultureInfo.InvariantCulture);
                case float f: return f.ToString("R", System.Globalization.CultureInfo.InvariantCulture);
                case bool b: return b ? "1" : "0";
                default: return Convert.ToString(v, System.Globalization.CultureInfo.InvariantCulture) ?? string.Empty;
            }
        }
    }
}