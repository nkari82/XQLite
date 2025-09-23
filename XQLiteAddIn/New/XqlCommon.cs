// XqlCommon.cs
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Runtime.InteropServices;


namespace XQLite.AddIn
{
    internal static class XqlCommon
    {
        public static void ReleaseCom(object? o)
        {
            try { if (o != null && Marshal.IsComObject(o)) Marshal.FinalReleaseComObject(o); } catch { }
        }

        public static string CreateTempDir(string prefix)
        {
            var root = Path.Combine(Path.GetTempPath(), prefix + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(root);
            return root;
        }

        public static void TryDeleteDir(string dir) 
        { 
            try 
            { 
                if (Directory.Exists(dir)) 
                    Directory.Delete(dir, true); 
            } 
            catch { } 
        }

        public static void SafeZipDirectory(string dir, string outZip)
        {
            try
            {
                if (File.Exists(outZip)) 
                    File.Delete(outZip);

                ZipFile.CreateFromDirectory(dir, outZip, CompressionLevel.Fastest, includeBaseDirectory: false);
            }
            catch { }
        }

        public static string CsvEscape(string s)
        {
            if (s == null) return "";
            bool needQuote = s.Contains(',') || s.Contains('"') || s.Contains('\n') || s.Contains('\r');
            return needQuote ? "\"" + s.Replace("\"", "\"\"") + "\"" : s;
        }

        public static string ValueToString(object? v)
        {
            if (v == null) return "";
            if (v is bool b) return b ? "TRUE" : "FALSE";
            if (v is DateTime dt) return dt.ToString("o");
            if (v is double d) return d.ToString("R", CultureInfo.InvariantCulture);
            if (v is float f) return ((double)f).ToString("R", CultureInfo.InvariantCulture);
            if (v is decimal m) return ((double)m).ToString("R", CultureInfo.InvariantCulture);
            return Convert.ToString(v, CultureInfo.InvariantCulture) ?? "";
        }

        public static IEnumerable<List<T>> Chunk<T>(IReadOnlyList<T> list, int size)
        {
            if (size <= 0) size = 1000;
            for (int i = 0; i < list.Count; i += size) 
                yield return list.Skip(i).Take(Math.Min(size, list.Count - i)).ToList();
        }

        public static void InterlockedMax(ref long target, long value)
        {
            while (true)
            {
                long cur = System.Threading.Volatile.Read(ref target);
                if (value <= cur) return;
                if (System.Threading.Interlocked.CompareExchange(ref target, value, cur) == cur) return;
            }
        }
    }
}