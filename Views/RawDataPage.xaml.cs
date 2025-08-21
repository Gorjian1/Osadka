using Osadka.ViewModels;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Text.RegularExpressions;

namespace Osadka.Views
{
    public partial class RawDataPage : UserControl
    {
        public RawDataPage(RawDataViewModel vm)
        {
            InitializeComponent();
            DataContext = vm;
        }

        private void DataGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

        }

        private void DataGrid_SelectionChanged_1(object sender, SelectionChangedEventArgs e)
        {

        }
        private void Control_DragOver(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                var files = (string[])e.Data.GetData(DataFormats.FileDrop);



                if (files.Any(f => f.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) ||
                                   f.EndsWith(".xlsm", StringComparison.OrdinalIgnoreCase)))
                {
                    e.Effects = DragDropEffects.Copy;
                }
                else
                    e.Effects = DragDropEffects.None;
            }
            else
                e.Effects = DragDropEffects.None;

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
            {
                vm.LoadWorkbookFromFile(path);
            }
        }


        // --- Allow digits, minus, dot and comma in limit textboxes ---
        private static readonly Regex _allowed = new Regex(@"^[0-9\.,\-]+$");

        private void Limit_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            e.Handled = !_allowed.IsMatch(e.Text);
        }

        private void Limit_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Space) e.Handled = true;
        }

        private void Limit_Pasting(object sender, DataObjectPastingEventArgs e)
        {
            if (e.DataObject.GetDataPresent(DataFormats.Text))
            {
                var text = (string)e.DataObject.GetData(DataFormats.Text);
                if (string.IsNullOrWhiteSpace(text) || !_allowed.IsMatch(text.Replace(" ", "")))
                    e.CancelCommand();
            }
            else
            {
                e.CancelCommand();
            }
        }

    }
}