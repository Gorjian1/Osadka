using System;
using System.Windows.Controls;
using System.Windows.Input;
using Osadka.ViewModels;

namespace Osadka.Views
{
    public partial class CycleGroupsPage : UserControl, IDisposable
    {
        private readonly CycleGroupsViewModel _viewModel;

        public CycleGroupsPage(CycleGroupsViewModel viewModel)
        {
            InitializeComponent();
            DataContext = _viewModel = viewModel;
        }

        public void Dispose()
        {
            _viewModel.Dispose();
        }

        private void OnRowClicked(object sender, MouseButtonEventArgs e)
        {
            if (sender is not Border border)
                return;

            if (border.DataContext is not CycleGroupRow row)
                return;

            row.IsHighlighted = !row.IsHighlighted;
            e.Handled = true;
        }
    }
}
