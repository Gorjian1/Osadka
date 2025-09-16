using AutoUpdaterDotNET;
using Osadka.Services;
using Osadka.ViewModels;
using System;
using System.ComponentModel;
using System.Threading.Tasks;
using System.Windows;

namespace Osadka
{
    public partial class MainWindow : Window
    {
        private readonly MainViewModel _vm;
        public string VersionInfo { get; }

        public MainWindow()
        {
            InitializeComponent();
            DataContext = this;
            _vm = new MainViewModel();
            DataContext = _vm;

            this.Closing += MainWindow_Closing;
            VersionInfo = $"Текущая версия: {UpdateService.CurrentVersionString}";
        }
        private void MainWindow_Closing(object? sender, CancelEventArgs e)
        {

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
