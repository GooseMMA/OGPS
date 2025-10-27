using System;
using System.Diagnostics;
using System.IO;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Media.Imaging;
using Microsoft.Win32;
using OGPSLauncher.Utils;

namespace OGPSLauncher
{
    public partial class MainWindow : Window
    {
        private string _gamePath = "";
        private const string GameExeName = "PixelWorlds.exe";
        private const string LauncherExeName = "OGPSLauncher.exe";
        private const string GameFolderName = "OGPSGame";

        public MainWindow()
        {
            InitializeComponent();
            LoadSavedPath();
            _ = LoadNewsAndPatchNotes(); // Загружаем данные при старте
        }

        private void LoadSavedPath()
        {
            _gamePath = Registry.GetValue(@"HKEY_CURRENT_USER\Software\OGPSLauncher", "InstallPath", "") as string ?? "";
            bool isInstalled = !string.IsNullOrEmpty(_gamePath) && File.Exists(Path.Combine(_gamePath, GameExeName));

            InstallButton.IsEnabled = !isInstalled;
            PlayButton.IsEnabled = isInstalled;
            MenuButton.IsEnabled = isInstalled;
            StatusText.Text = isInstalled 
                ? $"Установлено: {_gamePath}" 
                : "Игра не установлена";
        }

        private async void InstallButton_Click(object sender, RoutedEventArgs e)
        {
            using var dialog = new System.Windows.Forms.FolderBrowserDialog();
            dialog.Description = "Выберите папку для установки OGPS";
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                _gamePath = dialog.SelectedPath;
                StatusText.Text = "Установка...";

                try
                {
                    Directory.CreateDirectory(_gamePath);
                    Installer.InstallGame(_gamePath);
                    Installer.CopyLauncherToGameFolder(_gamePath, LauncherExeName);
                    ShortcutCreator.CreateShortcutOnDesktop(_gamePath, LauncherExeName, "OGPS");
                    Registry.SetValue(@"HKEY_CURRENT_USER\Software\OGPSLauncher", "InstallPath", _gamePath);

                    LoadSavedPath();
                    await LoadNewsAndPatchNotes();
                    StatusText.Text = "Установка завершена!";
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка: {ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                    StatusText.Text = "Ошибка установки";
                }
            }
        }

        private async void PlayButton_Click(object sender, RoutedEventArgs e)
        {
            string gameExe = Path.Combine(_gamePath, GameExeName);
            if (!File.Exists(gameExe))
            {
                MessageBox.Show("Файл PixelWorlds.exe не найден.", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            StatusText.Text = "Обновление модов...";
            await GitHubUpdater.UpdateModsAsync(_gamePath);

            StatusText.Text = "Запуск игры...";
            Process.Start(gameExe);
            Application.Current.Shutdown();
        }

        private void MenuButton_Click(object sender, RoutedEventArgs e)
        {
            OptionsMenu.IsOpen = true;
        }

        private void OpenFolderFromMenu_Click(object sender, RoutedEventArgs e)
        {
            OptionsMenu.IsOpen = false;
            if (!string.IsNullOrEmpty(_gamePath) && Directory.Exists(_gamePath))
            {
                Process.Start("explorer.exe", _gamePath);
            }
        }

        private void UninstallFromMenu_Click(object sender, RoutedEventArgs e)
        {
            OptionsMenu.IsOpen = false;
            if (string.IsNullOrEmpty(_gamePath) || !Directory.Exists(_gamePath))
            {
                MessageBox.Show("Игра не установлена.", "Информация", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var result = MessageBox.Show(
                "Вы уверены, что хотите удалить игру и все настройки?\nЭто действие нельзя отменить.",
                "Подтверждение удаления",
                MessageBoxButton.YesNo,
                MessageBoxImage.Warning
            );

            if (result == MessageBoxResult.Yes)
            {
                try
                {
                    Directory.Delete(_gamePath, true);
                    Registry.CurrentUser.DeleteSubKeyTree(@"Software\OGPSLauncher", false);
                    _gamePath = "";
                    LoadSavedPath();
                    StatusText.Text = "Игра и настройки удалены.";
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка при удалении: {ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private async Task LoadNewsAndPatchNotes()
        {
            try
            {
                using var client = new HttpClient();
                string newsUrl = "https://raw.githubusercontent.com/GooseMMA/OGPS/main/news.txt";
                string patchUrl = "https://raw.githubusercontent.com/GooseMMA/OGPS/main/patchnotes.txt";
                string screenshotUrl = "https://raw.githubusercontent.com/GooseMMA/OGPS/main/update_screenshot.png";

                NewsText.Text = await client.GetStringAsync(newsUrl);
                PatchNotesText.Text = await client.GetStringAsync(patchUrl);

                var imageBytes = await client.GetByteArrayAsync(screenshotUrl);
                var bitmap = new BitmapImage();
                using (var stream = new MemoryStream(imageBytes))
                {
                    bitmap.BeginInit();
                    bitmap.CacheOption = BitmapCacheOption.OnLoad;
                    bitmap.StreamSource = stream;
                    bitmap.EndInit();
                }
                UpdateScreenshot.Source = bitmap;

                UpdateStatusText.Text = "Проверка обновлений...";
                bool updated = await GitHubUpdater.UpdateModsAsync(_gamePath);
                UpdateStatusText.Text = updated ? "✅ Обновлено!" : "✔️ Актуально";
            }
            catch
            {
                NewsText.Text = "Не удалось загрузить новости.";
                PatchNotesText.Text = "Не удалось загрузить патчноуты.";
                UpdateStatusText.Text = "❌ Ошибка";
                UpdateScreenshot.Source = new BitmapImage(new Uri("pack://application:,,,/Resources/logo.png"));
            }
        }
    }
}