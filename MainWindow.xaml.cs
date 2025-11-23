using AutoUpdaterDotNET;
using Osadka.Services;
using System;
using System.ComponentModel;
using System.Windows;

namespace Osadka
{
    public partial class MainWindow : Window
    {
        public string VersionInfo { get; }

        public MainWindow()
        {
            InitializeComponent();

            // DataContext будет установлен в App.xaml.cs через DI

            this.Closing += MainWindow_Closing;
            VersionInfo = $"Текущая версия: {UpdateService.CurrentVersionString}";
        }
        private void MainWindow_Closing(object? sender, CancelEventArgs e)
        {
            // TODO: Можно добавить проверку несохраненных изменений
        }

        private void OnCheckUpdateClick(object sender, RoutedEventArgs e)
        {
            if (!Uri.TryCreate(UpdateService.ManifestUrl, UriKind.Absolute, out var uri) ||
                !(uri.Host.Equals("raw.githubusercontent.com", StringComparison.OrdinalIgnoreCase) ||
                  uri.Host.EndsWith("github.io", StringComparison.OrdinalIgnoreCase)))
            {
                MessageBox.Show(
                    $"Обнаружен недопустимый источник обновлений:\n{UpdateService.ManifestUrl}",
                    "Ошибка обновления",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
                return;
            }

            AutoUpdater.Start(UpdateService.ManifestUrl);
        }
    }
}
