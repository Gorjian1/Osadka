using CommunityToolkit.Mvvm.ComponentModel;
using System.ComponentModel;

namespace Osadka.Models
{
    public partial class GeneralReport : ObservableObject, IDataErrorInfo
    {
        [ObservableProperty] private string _MaxGeneralSettelmentNumbers = string.Empty;
        [ObservableProperty] private string _MinGeneralSettelmentNumbers = string.Empty;
        [ObservableProperty] private string _AvgGeneralSettelmentNumbers = string.Empty;

        [ObservableProperty] private string _MaxRelativeNumbers = string.Empty;
        [ObservableProperty] private string _MinRelativeNumbers = string.Empty;
        [ObservableProperty] private string _AvgRelativeNumbers = string.Empty;

        [ObservableProperty] private float _MaxGeneralValue;
        [ObservableProperty] private float _MinGeneralValue;
        [ObservableProperty] private float _AvgGeneralValue;

        [ObservableProperty] private float _MaxRelativeValue;
        [ObservableProperty] private float _MinRelativeValue;
        [ObservableProperty] private float _AvgRelativeValue;

        [ObservableProperty] private string _CycleLable = string.Empty;
        [ObservableProperty] private string _GeneralExtrema = "-";

        [ObservableProperty] private string _RelativeExtrema = "-";

        public string Error => null!;
        public string this[string column] => column switch { _ => string.Empty };
    }
}
