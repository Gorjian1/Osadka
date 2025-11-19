using Osadka.Models;

namespace Osadka.Services.Abstractions;

/// <summary>
/// Сервис для сохранения и загрузки проектов (.osd файлы)
/// </summary>
public interface IProjectService
{
    /// <summary>
    /// Загружает проект из файла
    /// </summary>
    /// <param name="path">Путь к файлу .osd</param>
    /// <returns>Данные проекта + путь к DWG (если есть)</returns>
    (ProjectData Data, string? DwgPath) Load(string path);

    /// <summary>
    /// Сохраняет проект в файл
    /// </summary>
    /// <param name="path">Путь для сохранения</param>
    /// <param name="data">Данные проекта</param>
    /// <param name="dwgPath">Путь к DWG файлу (опционально)</param>
    void Save(string path, ProjectData data, string? dwgPath = null);
}
