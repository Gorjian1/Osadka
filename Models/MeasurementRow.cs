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
}

