using System.Windows;

namespace Osadka.Services.Abstractions;

/// <summary>
/// Абстракция для показа MessageBox (для тестируемости ViewModels)
/// </summary>
public interface IMessageBoxService
{
    void Show(string message, string title = "");

    bool Confirm(string message, string title = "Подтверждение");

    MessageBoxResult ShowWithOptions(
        string message,
        string title,
        MessageBoxButton buttons,
        MessageBoxImage image = MessageBoxImage.None);
}
