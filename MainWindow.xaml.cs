using System;
using System.Threading.Tasks;
using System.Windows;
using Osadka.Services;

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
            var current = UpdateService.CurrentVersion;
            Version latest;

            try
            {
                latest = await UpdateService.GetLatestVersionAsync();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Не удалось проверить обновления:\n{ex.Message}",
                                "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            string msg;
            var icon = MessageBoxImage.Information;

            if (latest > current)
            {
                msg = $"Доступна новая версия: {latest}. Текущая: {current}.";
                icon = MessageBoxImage.Warning;
            }
            else
            {
                msg = $"Вы используете последнюю версию: {current}.";
            }

            MessageBox.Show(msg, "Проверка обновлений", MessageBoxButton.OK, icon);
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
