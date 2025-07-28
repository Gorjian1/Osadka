using System.Windows.Controls;
using Osadka.ViewModels;

namespace Osadka.Views
{
    public partial class DynamicsGrafficPage : UserControl
    {
        public DynamicsGrafficPage(DynamicsGrafficViewModel vm)
        {
            InitializeComponent();
            DataContext = vm;
            AttachModel(vm);

            vm.PropertyChanged += (_, e) =>
            {
                if (e.PropertyName == nameof(vm.PlotModel))
                    AttachModel(vm);
            };

            Unloaded += (_, __) => plotView.Model = null;
        }

        private void AttachModel(DynamicsGrafficViewModel vm)
        {
            if (vm.PlotModel != null && !ReferenceEquals(plotView.Model, vm.PlotModel))
                plotView.Model = vm.PlotModel;
        }
    }
}