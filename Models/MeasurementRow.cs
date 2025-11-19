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
    /// Проверяет, валидна ли точка для включения в расчёты.
    /// Точка невалидна если имеет текстовый статус или отсутствующее значение.
    /// </summary>
    public bool IsValidForCalculation()
    {
        // Если нет числового значения Total, точка невалидна
        if (Total is not double total || double.IsNaN(total) || double.IsInfinity(total))
            return false;

        // Если MarkRaw содержит текстовый статус (не число), точка невалидна
        if (!string.IsNullOrWhiteSpace(MarkRaw) && !Mark.HasValue)
            return false;

        // Если SettlRaw содержит статус "новая", это новая точка - невалидна для расчёта осадки
        if (!string.IsNullOrWhiteSpace(SettlRaw) &&
            SettlRaw.Contains("нов", System.StringComparison.OrdinalIgnoreCase))
            return false;

        // Если TotalRaw содержит "-" или другой текст, а Total null - невалидна
        if (!string.IsNullOrWhiteSpace(TotalRaw) && TotalRaw.Trim() == "-")
            return false;

        return true;
    }

    /// <summary>
    /// Проверяет валидность конкретного значения (Mark/Settl/Total)
    /// </summary>
    public bool IsValidValue(double? value, string rawValue)
    {
        if (value is not double v || double.IsNaN(v) || double.IsInfinity(v))
            return false;

        // Если raw содержит текстовый статус, невалидно
        if (!string.IsNullOrWhiteSpace(rawValue))
        {
            if (rawValue.Contains("нов", System.StringComparison.OrdinalIgnoreCase) ||
                rawValue.Contains("унич", System.StringComparison.OrdinalIgnoreCase) ||
                rawValue.Contains("нет доступ", System.StringComparison.OrdinalIgnoreCase) ||
                rawValue.Trim() == "-")
                return false;
        }

        return true;
    }
}

