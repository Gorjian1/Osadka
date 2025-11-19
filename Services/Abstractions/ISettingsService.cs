namespace Osadka.Services.Abstractions;

/// <summary>
/// Интерфейс для работы с пользовательскими настройками
/// </summary>
public interface ISettingsService
{
    string? TemplatePath { get; set; }

    void Load();
    void Save();
}
