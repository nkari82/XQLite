using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;

namespace XQLite.AddIn
{
    public static class XqlSecrets
    {
        private static string Dir => Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "XQLite");
        private static string PathKey => System.IO.Path.Combine(Dir, "secrets.bin");

        public static void SaveApiKey(string apiKey)
        {
            Directory.CreateDirectory(Dir);
            var plain = Encoding.UTF8.GetBytes(apiKey ?? string.Empty);
            var protectedBytes = ProtectedData.Protect(plain, optionalEntropy: null, scope: DataProtectionScope.CurrentUser);
            File.WriteAllBytes(PathKey, protectedBytes);
        }

        public static string LoadApiKey()
        {
            try
            {
                if (!File.Exists(PathKey)) return string.Empty;
                var bytes = File.ReadAllBytes(PathKey);
                var unprot = ProtectedData.Unprotect(bytes, optionalEntropy: null, scope: DataProtectionScope.CurrentUser);
                return Encoding.UTF8.GetString(unprot);
            }
            catch { return string.Empty; }
        }

        public static void Clear()
        {
            try { if (File.Exists(PathKey)) File.Delete(PathKey); } catch { }
        }
    }
}