using System;
using System.Collections.Generic;
using System.IO;

namespace XQLite.AddIn
{
    internal static class XqlOutbox
    {
        private static readonly object _gate = new();
        private static string Dir => Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "XQLite");
        private static string PathBox => System.IO.Path.Combine(Dir, "outbox.ndjson");

        internal sealed class Item
        {
            public string table { get; set; } = string.Empty;
            public Dictionary<string, object?> row { get; set; } = new();
            public DateTime at { get; set; } = DateTime.UtcNow;
        }

        public static void Append(string table, Dictionary<string, object?> row)
        {
            try
            {
                Directory.CreateDirectory(Dir);
                var it = new Item { table = table, row = row, at = DateTime.UtcNow };
                var json = XqlJson.Serialize(it);
                lock (_gate) File.AppendAllText(PathBox, json + "\n");
            }
            catch { }
        }

        public static IEnumerable<Item> ReadAllAndTruncate()
        {
            lock (_gate)
            {
                if (!File.Exists(PathBox)) yield break;
                string tmp = PathBox + ".tmp";
                try
                {
                    File.Move(PathBox, tmp);
                    foreach (var line in File.ReadLines(tmp))
                    {
                        if (string.IsNullOrWhiteSpace(line)) continue;
                        Item? it = null; try { it = XqlJson.Deserialize<Item>(line); } catch { }
                        if (it != null) yield return it;
                    }
                }
                finally { try { if (File.Exists(tmp)) File.Delete(tmp); } catch { } }
            }
        }
    }
}