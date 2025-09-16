using System;
using System.Globalization;
using System.Windows.Data;

namespace Osadka.Converters
{
    /// <summary>
    /// «Мягкое» преобразование строки в double.
    /// - Принимает и точку, и запятую.
    /// - "1." / "2," трактует как 1 / 2 (без ожидания потери фокуса).
    /// - Пустую строку и одиночный "-" считает незавершённым вводом -> источник не трогаем.
    /// </summary>
    public class RelaxedDoubleConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture) => value;

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            var s = value?.ToString()?.Trim() ?? string.Empty;
            if (s.Length == 0 || s == "-") return Binding.DoNothing; // незавершённый ввод

            s = s.Replace(',', '.');

            // Разрешим незавершённую дробь: "12." -> "12"
            if (s.EndsWith(".")) s = s.Substring(0, s.Length - 1);

            if (double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out double d))
                return d; // возвращаем double сразу по мере ввода

            return Binding.DoNothing;
        }
    }
}
