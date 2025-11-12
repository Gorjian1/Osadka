using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Osadka.Services.Reports;
using OxyPlot;
using OxyPlot.Axes;
using OxyPlot.Legends;
using OxyPlot.Series;
using System;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;

namespace Osadka.ViewModels
{
    public partial class DynamicsGrafficViewModel : ObservableObject
    {
        public ObservableCollection<DynamicSeries> Lines { get; } = new();

        [ObservableProperty] private PlotModel _plotModel = new();

        public IRelayCommand ExportCommand { get; }

        private readonly RawDataViewModel _raw;
        private readonly DynamicsReportService _svc;

        public DynamicsGrafficViewModel(RawDataViewModel raw,
                                        DynamicsReportService svc)
        {
            _raw = raw;
            _svc = svc;

            ExportCommand = new RelayCommand(ExportPng, () => PlotModel.Series.Any());
            _raw.PropertyChanged += (_, __) => Rebuild();
            _raw.ActiveFilterChanged += (_, __) => Rebuild();

            Rebuild();
        }
        private void Rebuild()
        {
            Lines.Clear();
            PlotModel = new PlotModel
            {
                Title = "Динамика осадок по маркам",
                IsLegendVisible = true,
            };

            var palette = OxyPalettes.HueDistinct(20);
            int colorIdx = 0;
            var activeCycles = _raw.GetActiveCyclesSnapshot();
            var series = _svc.Build(activeCycles)
                             .Select(s => new DynamicSeries
                             {
                                 Id = s.Id,
                                 Points = new ObservableCollection<(int, double)>(
                                             s.Points.Select(p => (p.Cycle, p.Mark)))
                             });

            foreach (var ds in series)
            {
                Lines.Add(ds);

                var ls = new LineSeries
                {
                    Title = ds.Id,
                    Color = palette.Colors[colorIdx++ % palette.Colors.Count],
                    MarkerSize = 0
                };

                foreach (var p in ds.Points.OrderBy(p => p.Item1))
                    ls.Points.Add(new OxyPlot.DataPoint(p.Item1, p.Item2));


                PlotModel.Series.Add(ls);
            }

            PlotModel.Axes.Add(new LinearAxis { Position = AxisPosition.Left, Title = "Отметка, м" });
            PlotModel.Axes.Add(new LinearAxis { Position = AxisPosition.Bottom, Title = "Цикл №", MinimumPadding = .1 });

            PlotModel.InvalidatePlot(true);
            ExportCommand.NotifyCanExecuteChanged();
        }

        private void ExportPng()
        {
            var dlg = new Microsoft.Win32.SaveFileDialog
            {
                Filter = "PNG|*.png",
                FileName = $"Dynamics_{DateTime.Now:yyyyMMdd_HHmm}.png"
            };
            if (dlg.ShowDialog() != true) return;

            try
            {
                var exporter = new OxyPlot.Wpf.PngExporter
                {
                    Width = 1000,
                    Height = 500,
                };
                using var fs = System.IO.File.Create(dlg.FileName);
                exporter.Export(PlotModel, fs);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"PNG-экспорт: {ex.Message}", "Ошибка",
                                MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }

    public sealed class DynamicSeries
    {
        public string Id { get; set; } = string.Empty;
        public ObservableCollection<(int Cycle, double Mark)> Points { get; set; } = new();
    }
}
