using Osadka.ViewModels;
using System;
using System.ComponentModel;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using System.Collections.Specialized;

namespace Osadka.Views
{
    public partial class RawDataPage : UserControl
    {
        public RawDataPage(RawDataViewModel vm)
        {
            InitializeComponent();
            DataContext = vm;

            // Синхронизация порядка в комбобоксах тулбара:
            //   Объекты — по возрастанию
            //   Циклы   — по убыванию (правый = самый поздний)
            SetupToolbarSorting(vm);

            // Коммит активного TextBox на Enter/клик — пусть остаётся, это не мешает
            AddHandler(Keyboard.PreviewKeyDownEvent, new KeyEventHandler(CommitOnEnter), true);
            AddHandler(Mouse.PreviewMouseDownEvent, new MouseButtonEventHandler(CommitOnMouseDown), true);
        }

        private void SetupToolbarSorting(RawDataViewModel vm)
        {
            // Объекты — по возрастанию
            var viewObjects = CollectionViewSource.GetDefaultView(vm.ObjectNumbers);
            if (viewObjects != null)
            {
                viewObjects.SortDescriptions.Clear();
                viewObjects.SortDescriptions.Add(new SortDescription(string.Empty, ListSortDirection.Ascending));
            }

            // Циклы — по Number убыванию (правый = самый поздний)
            var viewCycles = CollectionViewSource.GetDefaultView(vm.CycleItems);
            if (viewCycles != null)
            {
                viewCycles.SortDescriptions.Clear();
                viewCycles.SortDescriptions.Add(new SortDescription(nameof(RawDataViewModel.CycleDisplayItem.Number),
                                                                   ListSortDirection.Descending));
            }

            vm.ObjectNumbers.CollectionChanged += (_, __) => viewObjects?.Refresh();
            vm.CycleItems.CollectionChanged += (_, __) => viewCycles?.Refresh();
        }


        // Разрешаем: пусто | "-" | "-12" | "12" | "12." | "12.3" | ".5" | ",5"
        private static readonly Regex _numericPattern = new(@"^\s*-?\d*(?:[.,]\d*)?\s*$", RegexOptions.Compiled);

        private static bool ResultingTextIsValid(TextBox tb, string incoming)
        {
            var current = tb.Text ?? string.Empty;
            var selStart = tb.SelectionStart;
            var selLen = tb.SelectionLength;

            // Смоделируем итоговую строку после ввода/вставки
            string proposed = current.Remove(selStart, selLen).Insert(selStart, incoming);
            return _numericPattern.IsMatch(proposed);
        }

        private void CommitOnEnter(object? sender, KeyEventArgs e)
        {
            if (e.Key != Key.Enter) return;
            CommitActiveTextBox();
        }

        private void CommitOnMouseDown(object? sender, MouseButtonEventArgs e)
        {
            CommitActiveTextBox();
        }

        private void CommitActiveTextBox()
        {
            if (FocusManager.GetFocusedElement(this) is TextBox tb)
            {
                BindingExpression? be = tb.GetBindingExpression(TextBox.TextProperty);
                be?.UpdateSource();
            }
        }

        // === Ограничители ввода ===

        private void Limit_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            var tb = (TextBox)sender;
            // Проверяем будущую строку целиком, а не одиночный символ.
            e.Handled = !ResultingTextIsValid(tb, e.Text);
        }

        private void Limit_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Space)
            {
                e.Handled = true;
                return;
            }

            // Разрешаем системные клавиши без проверки текста
            if (e.Key == Key.Back || e.Key == Key.Delete ||
                e.Key == Key.Left || e.Key == Key.Right ||
                e.Key == Key.Home || e.Key == Key.End)
            {
                e.Handled = false;
            }
        }

        private void Limit_Pasting(object sender, DataObjectPastingEventArgs e)
        {
            if (!e.DataObject.GetDataPresent(DataFormats.Text))
            {
                e.CancelCommand();
                return;
            }

            var text = (string)e.DataObject.GetData(DataFormats.Text) ?? string.Empty;
            var tb = (TextBox)sender;

            if (!ResultingTextIsValid(tb, text))
                e.CancelCommand();
        }

        // === DnD Excel оставляем как было ===

        private void Control_DragOver(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                var files = (string[])e.Data.GetData(DataFormats.FileDrop);
                if (files.Any(f => f.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) ||
                                   f.EndsWith(".xlsm", StringComparison.OrdinalIgnoreCase)))
                    e.Effects = DragDropEffects.Copy;
                else
                    e.Effects = DragDropEffects.None;
            }
            else
            {
                e.Effects = DragDropEffects.None;
            }

            e.Handled = true;
        }

        private void Control_Drop(object sender, DragEventArgs e)
        {
            if (!e.Data.GetDataPresent(DataFormats.FileDrop)) return;

            var files = (string[])e.Data.GetData(DataFormats.FileDrop);
            var path = files.FirstOrDefault(f => f.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) ||
                                                 f.EndsWith(".xlsm", StringComparison.OrdinalIgnoreCase));
            if (path == null) return;

            if (DataContext is RawDataViewModel vm)
                vm.ImportFromExcel(path);
        }
    }
}
