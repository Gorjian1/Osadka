using System;
using System.Windows.Controls;
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
    }
}
