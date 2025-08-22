using CommunityToolkit.Mvvm.ComponentModel;
using System.ComponentModel;

namespace Osadka.Models;

public partial class MeasurementRow : ObservableObject
{
    [ObservableProperty] private string _id = string.Empty;

    // --- СТАРЫЕ поля (оставлены для совместимости UI/шаблонов, можно убрать позже) ---
    [ObservableProperty] private string _markRaw = string.Empty;
    [ObservableProperty] private string _settlRaw = string.Empty;
    [ObservableProperty] private string _totalRaw = string.Empty;
    [ObservableProperty] private double? _mark;
    [ObservableProperty] private double? _settl;
    [ObservableProperty] private double? _total;

    // --- НОВЫЕ поля (единый источник истины) ---
    [ObservableProperty] private double? _x;
    [ObservableProperty] private double? _y;
    [ObservableProperty] private double? _h;

    [ObservableProperty] private double? _dx;
    [ObservableProperty] private double? _dy;
    [ObservableProperty] private double? _dh;
    [ObservableProperty] private double? _vector;

    [ObservableProperty] private int _cycle;

    // Форматы отображения (используйте в таблицах, если нужно):
    public string DxDisplay => Dx is double v ? v.ToString("F1") : string.Empty;
    public string DyDisplay => Dy is double v ? v.ToString("F1") : string.Empty;
    public string DhDisplay => Dh is double v ? v.ToString("F1") : string.Empty;
    public string VectorDisplay => Vector is double v ? v.ToString("F1") : string.Empty;

    // Старые дисплеи оставляем (не используйте их для новой логики)
    public string MarkDisplay => Mark is double v ? v.ToString("F3") : MarkRaw;
    public string SettlDisplay => Settl is double v ? v.ToString("F1") : SettlRaw;
    public string TotalDisplay => Total is double v ? v.ToString("F1") : TotalRaw;
}
