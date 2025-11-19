namespace Osadka.Services.Abstractions;

/// <summary>
/// Абстракция для файловых диалогов (OpenFileDialog, SaveFileDialog)
/// </summary>
public interface IFileDialogService
{
    /// <summary>
    /// Показывает диалог открытия файла
    /// </summary>
    /// <param name="filter">Фильтр файлов (например, "Excel Files (*.xlsx)|*.xlsx")</param>
    /// <param name="initialDirectory">Начальная директория</param>
    /// <returns>Путь к выбранному файлу или null, если отменено</returns>
    string? OpenFile(string filter, string? initialDirectory = null);

    /// <summary>
    /// Показывает диалог сохранения файла
    /// </summary>
    /// <param name="filter">Фильтр файлов</param>
    /// <param name="defaultFileName">Имя файла по умолчанию</param>
    /// <returns>Путь для сохранения или null, если отменено</returns>
    string? SaveFile(string filter, string? defaultFileName = null);
}
