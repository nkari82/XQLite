using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace XQLite.AddIn
{
#if !NET5_0_OR_GREATER && !NETCOREAPP3_0_OR_GREATER
    internal static class Compat
    {
        internal static TValue? GetValueOrDefault<TKey, TValue>(
            this IDictionary<TKey, TValue> dict, TKey key)
        {
            return dict.TryGetValue(key, out var value) ? value : default;
        }

        internal static TValue GetValueOrDefault<TKey, TValue>(
            this IDictionary<TKey, TValue> dict, TKey key, TValue defaultValue)
        {
            return dict.TryGetValue(key, out var value) ? value : defaultValue;
        }

        internal static TValue? GetValueOrDefault<TKey, TValue>(
            this Dictionary<TKey, TValue> dict, TKey key)
        {
            return dict.TryGetValue(key, out var value) ? value : default;
        }

        internal static TValue GetValueOrDefault<TKey, TValue>(
            this Dictionary<TKey, TValue> dict, TKey key, TValue defaultValue)
        {
            return dict.TryGetValue(key, out var value) ? value : defaultValue;
        }

        internal static string ToHexString(byte[] bytes)
        {
            if (bytes == null) throw new ArgumentNullException(nameof(bytes));
            // 빠른 대문자 HEX 변환 (할당 1회)
            var chars = new char[bytes.Length * 2];
            int i = 0;
            foreach (var b in bytes)
            {
                int hi = b >> 4, lo = b & 0xF;
                chars[i++] = (char)(hi > 9 ? ('A' + hi - 10) : ('0' + hi));
                chars[i++] = (char)(lo > 9 ? ('A' + lo - 10) : ('0' + lo));
            }
            return new string(chars);
        }

        internal static void Deconstruct<TKey, TValue>(this KeyValuePair<TKey, TValue> kvp,
                                             out TKey key, out TValue value)
        {
            key = kvp.Key;
            value = kvp.Value;
        }

        internal static async Task WriteLineAsync(this StreamWriter sw, string str, CancellationToken ct)
        {
            ct.ThrowIfCancellationRequested();
            await sw.WriteLineAsync(str);
            ct.ThrowIfCancellationRequested();
        }

        internal static int Clamp(int value, int min, int max)
            => value < min ? min : (value > max ? max : value);

        internal static long Clamp(long value, long min, long max)
            => value < min ? min : (value > max ? max : value);

        internal static double Clamp(double value, double min, double max)
            => value < min ? min : (value > max ? max : value);
    }
#endif
}
