using System;
using System.Globalization;
using System.Windows.Data;

namespace Osadka.Converters;

public class RelaxedDoubleConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)  => value;

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
    {
        var s = value?.ToString()?.Trim() ?? "";
        s = s.Replace(',', '.');

        return double.TryParse(s, NumberStyles.Any,
                               CultureInfo.InvariantCulture, out double d)
               ? (double?)d
               : Binding.DoNothing;
    }
}
