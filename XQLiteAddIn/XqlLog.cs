// XqlLog.cs
using System.Diagnostics;

namespace XQLite.AddIn
{
    internal static class XqlLog
    {
        public static void Info(string msg, string table = "*")
        {
            Debug.WriteLine($"[XQL] {msg}");
            try { XqlFileLogger.Write("INFO", table, msg); } catch { }
        }
        public static void Warn(string msg, string table = "*")
        {
            Debug.WriteLine($"[XQL][WARN] {msg}");
            try { XqlFileLogger.Write("WARN", table, msg); } catch { }
        }
        public static void Error(string msg, string table = "*")
        {
            Debug.WriteLine($"[XQL][ERR] {msg}");
            try { XqlFileLogger.Write("ERR", table, msg); } catch { }
        }
    }
}
