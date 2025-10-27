using System;
using System.IO;
using System.IO.Compression;
using System.Net.Http;
using System.Threading.Tasks;

namespace OGPSLauncher.Utils
{
    public static class GitHubUpdater
    {
        private static readonly HttpClient client = new();
        private const string VersionUrl = "https://raw.githubusercontent.com/GooseMMA/OGPS/main/mods_version.txt";
        private const string ModsZipUrl = "https://github.com/GooseMMA/OGPS/releases/download/Mods/mods.zip";

        public static async Task<bool> UpdateModsAsync(string gamePath)
        {
            try
            {
                string remoteVersion = (await client.GetStringAsync(VersionUrl)).Trim();
                string modsDir = Path.Combine(gamePath, "Mods");
                string versionFile = Path.Combine(modsDir, ".launcher_version");
                string localVersion = File.Exists(versionFile) ? await File.ReadAllTextAsync(versionFile) : "";

                if (remoteVersion == localVersion)
                    return false;

                byte[] zipData = await client.GetByteArrayAsync(ModsZipUrl);

                if (Directory.Exists(modsDir))
                    Directory.Delete(modsDir, true);
                Directory.CreateDirectory(modsDir);

                string tempZip = Path.Combine(Path.GetTempPath(), "mods_temp.zip");
                await File.WriteAllBytesAsync(tempZip, zipData);
                ZipFile.ExtractToDirectory(tempZip, modsDir);
                File.Delete(tempZip);

                await File.WriteAllTextAsync(versionFile, remoteVersion);
                return true;
            }
            catch
            {
                return false;
            }
        }
    }
}