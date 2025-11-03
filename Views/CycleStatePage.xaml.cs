using System.Windows.Controls;
using Osadka.ViewModels;

namespace Osadka.Views
{
    public partial class CycleStatePage : UserControl
    {
        public CycleStatePage(CycleStateViewModel viewModel)
        {
            InitializeComponent();
            DataContext = viewModel;
        }
    }
}
