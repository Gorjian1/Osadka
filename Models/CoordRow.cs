using CommunityToolkit.Mvvm.ComponentModel;
using System.ComponentModel;

namespace Osadka.Models;

public partial class CoordRow : ObservableObject, IDataErrorInfo
{
    [ObservableProperty] private double _X;
    [ObservableProperty] private double _Y;

    public string Error => null!;

    public string this[string column] => column switch
    {
        _ => string.Empty
    };
}
