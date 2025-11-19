namespace Osadka.Services.Abstractions;

/// <summary>
/// Абстракция для работы с файлами (Process.Start и т.д.)
/// </summary>
public interface IFileService
{
    /// <summary>
    /// Открывает файл в приложении по умолчанию (Process.Start)
    /// </summary>
    void OpenInDefaultApp(string path);

    /// <summary>
    /// Проверяет существование файла
    /// </summary>
    bool FileExists(string path);
}
