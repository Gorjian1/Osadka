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

            // Используем UpdateService для версии
            VersionInfo = $"Текущая версия: {UpdateService.CurrentVersionString}";
        }

        private async void OnCheckUpdateClick(object sender, RoutedEventArgs e)
        {
            AutoUpdater.Start(UpdateService.ManifestUrl);

        }// В файле MainWindow.xaml.cs, внутри класса MainWindow
        private async void OnTestReportClick(object sender, RoutedEventArgs e)
        {
            try
            {
                // Отправляем тестовое сообщение
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
