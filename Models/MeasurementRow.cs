using CommunityToolkit.Mvvm.ComponentModel;
using System.ComponentModel;

namespace Osadka.Models;

public partial class MeasurementRow : ObservableObject
{
    [ObservableProperty] private string _id = string.Empty;

    [ObservableProperty] private string _markRaw = string.Empty;
    [ObservableProperty] private string _settlRaw = string.Empty;
    [ObservableProperty] private string _totalRaw = string.Empty;

    [ObservableProperty] private double? _mark;
    [ObservableProperty] private double? _settl;
    [ObservableProperty] private double? _total;

    [ObservableProperty] private int _cycle;

    public string MarkDisplay => Mark is double v ? v.ToString("F3") : MarkRaw;
    public string SettlDisplay => Settl is double v ? v.ToString("F1") : SettlRaw;
    public string TotalDisplay => Total is double v ? v.ToString("F1") : TotalRaw;

    /// <summary>
    /// Определяет, должна ли точка участвовать в вычислениях для текущего цикла.
    /// Точка исключается если имеет статус "Нет доступа" или "Уничтожена".
    /// </summary>
    public bool IsAvailableForCalculations()
    {
        string combined = string.Join(" ", new[] { MarkRaw, SettlRaw, TotalRaw })
            .Trim()
            .ToLowerInvariant();

        // Исключаем "Нет доступа"
        if (combined.Contains("нет") &&
            (combined.Contains("доступ") || combined.Contains("наблю") || combined.Contains("изм")))
        {
            return false;
        }

        // Исключаем "Уничтожена"
        if (combined.Contains("уничт") || combined.Contains("снес") ||
            combined.Contains("демонт") || combined.Contains("разруш"))
        {
            return false;
        }

        return true;
    }
}

