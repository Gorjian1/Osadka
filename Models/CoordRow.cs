using CommunityToolkit.Mvvm.ComponentModel;
using System.ComponentModel;

namespace Osadka.Models;

public partial class CoordRow : ObservableObject, IDataErrorInfo
{
    [ObservableProperty] private double _X;
    [ObservableProperty] private double _Y;
    [ObservableProperty] private double? _H; 

    public string Error => null!;

    public string this[string column] => column switch
    {
        _ => string.Empty
    };
}
