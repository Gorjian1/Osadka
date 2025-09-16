using CommunityToolkit.Mvvm.ComponentModel;
using System.ComponentModel;

public partial class CycleHeader : ObservableObject, IDataErrorInfo
{
    [ObservableProperty] private int _cycleNumber; 
    [ObservableProperty] private int _objectNumber;
    [ObservableProperty] private int _totalCycles; 
    [ObservableProperty] private int _totalObjects; 

    [ObservableProperty] private double? _maxNomen;
    [ObservableProperty] private double? _maxCalculated;
    [ObservableProperty] private double? _relNomen;
    [ObservableProperty] private double? _relCalculated;

    [ObservableProperty] private bool _isCycleEditable = true;

    public string Error => null!;
    public string this[string col] => col switch
    {
        nameof(CycleNumber) when CycleNumber <= 0 => "Номер цикла > 0",
        _ => string.Empty
    };
}
