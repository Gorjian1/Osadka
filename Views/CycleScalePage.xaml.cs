using System.Windows.Controls;
using Osadka.ViewModels;

namespace Osadka.Views
{
    public partial class CycleScalePage : Page
    {
        public CycleScalePage(CycleScaleViewModel viewModel)
        {
            InitializeComponent();
            DataContext = viewModel;
        }
    }
}
