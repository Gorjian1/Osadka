using System;
using System.Globalization;
using System.Linq;
using System.Windows.Data;

namespace Osadka.Converters;

public class JoinConverter : IValueConverter
{
    public string Separator { get; set; } = ", ";
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture) =>
        value switch
        {
            null => "",
            IEnumerable<string> list => string.Join(", ", list),
            var v when v is string s && s.Length > 0 => s,
            _ => ""
        };

    public object ConvertBack(object v, Type t, object p, CultureInfo c) => Binding.DoNothing;
}
