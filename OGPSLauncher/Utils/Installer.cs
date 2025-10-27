using System;
using System.IO;
using System.IO.Compression;
using System.Reflection;

namespace OGPSLauncher.Utils
{
    public static class Installer
    {
        public static void InstallGame(string installPath)
        {
            var assembly = Assembly.GetExecutingAssembly();
            using var zipStream = assembly.GetManifestResourceStream("OGPSLauncher.Resources.game.zip");
            if (zipStream == null)
                throw new FileNotFoundException("Архив game.zip не найден!");

            string tempZip = Path.Combine(Path.GetTempPath(), "ogps_game_temp.zip");
            using (var fs = File.Create(tempZip))
                zipStream.CopyTo(fs);

            ZipFile.ExtractToDirectory(tempZip, installPath, true);
            File.Delete(tempZip);
        }

        public static void CopyLauncherToGameFolder(string installPath, string launcherName)
        {
            string currentExe = Environment.ProcessPath!;
            string targetExe = Path.Combine(installPath, launcherName);
            File.Copy(currentExe, targetExe, true);
        }
    }
}