using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;


namespace XQLite.AddIn
{
    public static class XqlHash
    {
        [ThreadStatic] private static SHA256? _sha;
        public static string RowHash(IReadOnlyDictionary<string, object?> row)
        {
            _sha ??= SHA256.Create();
            var sb = new StringBuilder(512);
            foreach (var kv in row.OrderBy(k => k.Key, StringComparer.Ordinal))
            {
                sb.Append(kv.Key).Append('=');
                if (kv.Value is null) sb.Append('∅');
                else sb.Append(kv.Value is double d ? d.ToString("R") : kv.Value.ToString());
                sb.Append('|');
            }
            var bytes = Encoding.UTF8.GetBytes(sb.ToString());
            var hash = _sha.ComputeHash(bytes);
            return Compat.ToHexString(hash);
        }
    }
}
