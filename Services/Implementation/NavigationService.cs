using Osadka.Services.Abstractions;
using System;
using System.Collections.Generic;

namespace Osadka.Services.Implementation;

/// <summary>
/// Реализация сервиса навигации между страницами
/// </summary>
public class NavigationService : INavigationService
{
    private readonly Dictionary<string, Func<object>> _pageFactories = new();
    private readonly Dictionary<string, object> _cachedPages = new();
    private object? _currentPage;

    public object? CurrentPage
    {
        get => _currentPage;
        private set
        {
            if (_currentPage != value)
            {
                _currentPage = value;
                CurrentPageChanged?.Invoke(this, EventArgs.Empty);
            }
        }
    }

    public event EventHandler? CurrentPageChanged;

    public void RegisterPage(string pageKey, Func<object> factory)
    {
        _pageFactories[pageKey] = factory ?? throw new ArgumentNullException(nameof(factory));
    }

    public void NavigateTo(string pageKey)
    {
        if (string.IsNullOrWhiteSpace(pageKey))
        {
            return;
        }

        if (!_pageFactories.TryGetValue(pageKey, out var factory))
        {
            throw new InvalidOperationException($"Страница с ключом '{pageKey}' не зарегистрирована");
        }

        // Проверяем кеш (для страниц которые не нужно пересоздавать)
        // Пока создаем каждый раз новую, кроме Coord
        if (pageKey == "Coordinates")
        {
            if (!_cachedPages.TryGetValue(pageKey, out var cachedPage))
            {
                cachedPage = factory();
                _cachedPages[pageKey] = cachedPage;
            }
            CurrentPage = cachedPage;
        }
        else
        {
            // Создаем новую страницу каждый раз
            CurrentPage = factory();
        }
    }
}
