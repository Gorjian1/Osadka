using System.Windows;
using System.Windows.Controls;

namespace Osadka.UI
{
    // Публичный статический класс с прикреплёнными свойствами для генерации колонок
    public static class GridHelpers
    {
        public static readonly DependencyProperty ColumnsCountProperty =
            DependencyProperty.RegisterAttached(
                "ColumnsCount",
                typeof(int),
                typeof(GridHelpers),
                new PropertyMetadata(0, OnColumnsChanged));

        public static void SetColumnsCount(Grid grid, int value) => grid.SetValue(ColumnsCountProperty, value);
        public static int GetColumnsCount(Grid grid) => (int)grid.GetValue(ColumnsCountProperty);

        public static readonly DependencyProperty SharedGroupPrefixProperty =
            DependencyProperty.RegisterAttached(
                "SharedGroupPrefix",
                typeof(string),
                typeof(GridHelpers),
                new PropertyMetadata(null, OnColumnsChanged));

        public static void SetSharedGroupPrefix(Grid grid, string? value) => grid.SetValue(SharedGroupPrefixProperty, value);
        public static string? GetSharedGroupPrefix(Grid grid) => (string?)grid.GetValue(SharedGroupPrefixProperty);

        private static void OnColumnsChanged(DependencyObject d, DependencyPropertyChangedEventArgs _)
        {
            if (d is not Grid grid) return;

            int count = GetColumnsCount(grid);
            string? prefix = GetSharedGroupPrefix(grid);

            grid.ColumnDefinitions.Clear();
            if (count <= 0) return;

            for (int i = 0; i < count; i++)
            {
                var col = new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) };
                if (!string.IsNullOrWhiteSpace(prefix))
                    col.SharedSizeGroup = $"{prefix}{i}";
                grid.ColumnDefinitions.Add(col);
            }
        }
    }
}
