namespace Osadka.Services.Abstractions;

/// <summary>
/// Сервис для управления навигацией между страницами
/// </summary>
public interface INavigationService
{
    /// <summary>
    /// Текущая отображаемая страница (View)
    /// </summary>
    object? CurrentPage { get; }

    /// <summary>
    /// Событие изменения текущей страницы
    /// </summary>
    event EventHandler? CurrentPageChanged;

    /// <summary>
    /// Переход на страницу по ключу
    /// </summary>
    /// <param name="pageKey">Ключ страницы (например, "RawData", "SettlementDiff")</param>
    void NavigateTo(string pageKey);

    /// <summary>
    /// Регистрация фабрики для создания страницы
    /// </summary>
    /// <param name="pageKey">Ключ страницы</param>
    /// <param name="factory">Фабрика для создания View</param>
    void RegisterPage(string pageKey, Func<object> factory);
}
