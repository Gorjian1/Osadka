using System;
using System.Globalization;
using System.Windows.Data;

namespace Osadka.Converters
{
    /// <summary>
    /// Возвращает входное целое значение + 1 (для нумерации с 1).
    /// Используется с ItemsControl.AlternationIndex.
    /// </summary>
    public class AddOneConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is int i) return i + 1;
            if (value is string s && int.TryParse(s, out var j)) return j + 1;
            return value;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
            => Binding.DoNothing;
    }
}
