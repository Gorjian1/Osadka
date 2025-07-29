using AutoUpdaterDotNET;
using Osadka.Services;
using System;
using System.Threading.Tasks;
using System.Windows;

namespace Osadka
{
    public partial class MainWindow : Window
    {
        public string VersionInfo { get; }

        public MainWindow()
        {
            InitializeComponent();
            DataContext = this;

            VersionInfo = $"Текущая версия: {UpdateService.CurrentVersionString}";
        }

        private async void OnCheckUpdateClick(object sender, RoutedEventArgs e)
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

        private async void OnTestReportClick(object sender, RoutedEventArgs e)
        {
            try
            {
                await TelegramReporter.SendAsync("Тестовое сообщение: отчёт из Osadka ✔️");
                MessageBox.Show(
                    "Тестовое сообщение отправлено.",
                    "Успех",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information
                );
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"Ошибка при отправке тестового сообщения:\n{ex.Message}",
                    "Ошибка",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error
                );
            }
        }

    }
}
