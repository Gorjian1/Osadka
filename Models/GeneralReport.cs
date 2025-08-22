using CommunityToolkit.Mvvm.ComponentModel;
using System.ComponentModel;


namespace Osadka.Models
{
    public partial class GeneralReport : ObservableObject, IDataErrorInfo
    {
        [ObservableProperty] private string _MaxVectorIds = string.Empty;
        [ObservableProperty] private string _MinVectorIds = string.Empty;
        [ObservableProperty] private float _MaxVectorValue;
        [ObservableProperty] private float _MinVectorValue;

        [ObservableProperty] private string _MaxDxIds = string.Empty;
        [ObservableProperty] private string _MinDxIds = string.Empty;
        [ObservableProperty] private float _MaxDxValue;
        [ObservableProperty] private float _MinDxValue;

        [ObservableProperty] private string _MaxDyIds = string.Empty;
        [ObservableProperty] private string _MinDyIds = string.Empty;
        [ObservableProperty] private float _MaxDyValue;
        [ObservableProperty] private float _MinDyValue;

        [ObservableProperty] private string _MaxDhIds = string.Empty;
        [ObservableProperty] private string _MinDhIds = string.Empty;
        [ObservableProperty] private float _MaxDhValue;
        [ObservableProperty] private float _MinDhValue;


        [ObservableProperty] private string _CycleLable;

        public string Error => null!;

        public string this[string column] => column switch
        {
            _ => string.Empty
        };
    }
}
