using System;
using System.Globalization;
using System.Windows.Data;
using Osadka.ViewModels;

namespace Osadka.Converters
{
    /// <summary>
    /// На вход даём: [0] — индекс строки (AlternationIndex), [1] — RawDataViewModel.
    /// Возвращает "№ <Id>" если Id есть в DataRows, иначе "#<index+1>".
    /// </summary>
    public sealed class IndexToMarkLabelConverter : IMultiValueConverter
    {
        public object Convert(object[] values, Type targetType, object parameter, CultureInfo culture)
        {
            int idx = (values != null && values.Length > 0 && values[0] is int i) ? i : -1;
            var raw = (values != null && values.Length > 1) ? values[1] as RawDataViewModel : null;

            if (idx >= 0 && raw != null && idx < raw.DataRows.Count)
            {
                // Id — это «номер марки» из RawData
                // (см. заполнение Id при чтении Excel и коллекцию DataRows)
                // raw.DataRows[idx].Id -> строка с номером марки
                var id = raw.DataRows[idx].Id;
                return string.IsNullOrWhiteSpace(id) ? $"#{idx + 1}" : $"№ {id}";
            }

            return $"#{idx + 1}";
        }

        public object[] ConvertBack(object value, Type[] targetTypes, object parameter, CultureInfo culture)
            => throw new NotSupportedException();
    }
}
