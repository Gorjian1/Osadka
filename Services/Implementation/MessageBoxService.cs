using Osadka.Services.Abstractions;
using System.Windows;

namespace Osadka.Services.Implementation;

public class MessageBoxService : IMessageBoxService
{
    public void Show(string message, string title = "")
    {
        MessageBox.Show(message, title, MessageBoxButton.OK, MessageBoxImage.Information);
    }

    public bool Confirm(string message, string title = "Подтверждение")
    {
        var result = MessageBox.Show(message, title, MessageBoxButton.YesNo, MessageBoxImage.Question);
        return result == MessageBoxResult.Yes;
    }

    public MessageBoxResult ShowWithOptions(
        string message,
        string title,
        MessageBoxButton buttons,
        MessageBoxImage image = MessageBoxImage.None)
    {
        return MessageBox.Show(message, title, buttons, image);
    }
}
