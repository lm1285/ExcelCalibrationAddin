using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;

namespace ExcelCalibrationAddin.Vsto
{
    internal sealed class CloudSessionStore
    {
        private readonly string _path = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "ExcelCalibrationAddin",
            "cloud-session.dat");

        public string LoadToken()
        {
            try
            {
                if (!File.Exists(_path)) return string.Empty;
                var encrypted = File.ReadAllBytes(_path);
                var plain = ProtectedData.Unprotect(encrypted, null, DataProtectionScope.CurrentUser);
                return Encoding.UTF8.GetString(plain);
            }
            catch
            {
                return string.Empty;
            }
        }

        public void SaveToken(string token)
        {
            var directory = Path.GetDirectoryName(_path);
            if (!string.IsNullOrWhiteSpace(directory)) Directory.CreateDirectory(directory);
            var plain = Encoding.UTF8.GetBytes(token ?? string.Empty);
            var encrypted = ProtectedData.Protect(plain, null, DataProtectionScope.CurrentUser);
            File.WriteAllBytes(_path, encrypted);
        }

        public void Clear()
        {
            try { if (File.Exists(_path)) File.Delete(_path); } catch { }
        }
    }
}
