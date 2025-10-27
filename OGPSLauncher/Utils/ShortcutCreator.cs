using System;

namespace OGPSLauncher.Utils
{
    public static class ShortcutCreator
    {
        public static void CreateShortcutOnDesktop(string targetFolder, string targetExe, string shortcutName)
        {
            try
            {
                string desktop = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                string shortcutPath = Path.Combine(desktop, $"{shortcutName}.lnk");

                Type shellType = Type.GetTypeFromProgID("WScript.Shell");
                dynamic shell = Activator.CreateInstance(shellType);
                dynamic shortcut = shell.CreateShortcut(shortcutPath);

                shortcut.TargetPath = Path.Combine(targetFolder, targetExe);
                shortcut.WorkingDirectory = targetFolder;
                shortcut.Description = shortcutName;
                shortcut.Save();
            }
            catch { /* ignore */ }
        }
    }
}