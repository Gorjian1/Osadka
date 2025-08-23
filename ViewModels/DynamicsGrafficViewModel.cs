using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using OxyPlot;
using OxyPlot.Axes;
using OxyPlot.Series;
using System;
using System.Collections.ObjectModel;
using System.IO;
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

        // Сохраняем сигнатуру конструктора для DI, но сервис больше не требуется
        public DynamicsGrafficViewModel(RawDataViewModel raw, object? _unusedService = null)
        {
            _raw = raw;

            ExportCommand = new RelayCommand(ExportPng, () => PlotModel?.Series?.Any() == true);

            // Любое изменение данных (цикл/объект/вставка) — перестраиваем графики
            _raw.PropertyChanged += (_, __) => Rebuild();

            Rebuild();
        }

        private void Rebuild()
        {
            Lines.Clear();

            var pm = new PlotModel
            {
                Title = "Динамика вектора по маркам",
                IsLegendVisible = true,
            };

            var palette = OxyPlot.OxyPalettes.HueDistinct(24);
            int colorIdx = 0;

            // Берём текущий объект: цикл -> список строк измерений
            var cycles = _raw.CurrentCycles;
            if (cycles != null && cycles.Count > 0)
            {
                // Все марки, встречающиеся в циклах
                var ids = cycles.Values
                                .SelectMany(rows => rows.Select(r => r.Id))
                                .Distinct()
                                .OrderBy(id => id);

                foreach (var id in ids)
                {
                    // Собираем точки (цикл, Vector), пропуская пустые значения
                    var pts = cycles
                        .OrderBy(kv => kv.Key)
                        .Select(kv =>
                        {
                            var row = kv.Value.FirstOrDefault(r => r.Id == id);
                            return (Cycle: kv.Key, Val: row?.Vector);
                        })
                        .Where(t => t.Val.HasValue)
                        .Select(t => (t.Cycle, t.Val!.Value))
                        .ToList();

                    if (pts.Count == 0)
                        continue;

                    Lines.Add(new DynamicSeries
                    {
                        Id = id,
                        Points = new ObservableCollection<(int Cycle, double Value)>(pts)
                    });

                    var line = new LineSeries
                    {
                        Title = id,
                        Color = palette.Colors[colorIdx++ % palette.Colors.Count],
                        MarkerSize = 0
                    };

                    foreach (var p in pts)
                        line.Points.Add(new DataPoint(p.Cycle, p.Value));

                    pm.Series.Add(line);
                }
            }

            pm.Axes.Add(new LinearAxis
            {
                Position = AxisPosition.Left,
                Title = "Вектор, мм"
            });

            pm.Axes.Add(new LinearAxis
            {
                Position = AxisPosition.Bottom,
                Title = "Цикл №",
                MinimumPadding = .1,
                AbsoluteMinimum = 0
            });

            PlotModel = pm;
            ExportCommand.NotifyCanExecuteChanged();
        }

        private void ExportPng()
        {
            var dlg = new Microsoft.Win32.SaveFileDialog
            {
                Filter = "PNG|*.png",
                FileName = $"Dynamics_Vector_{DateTime.Now:yyyyMMdd_HHmm}.png"
            };
            if (dlg.ShowDialog() != true) return;

            try
            {
                var exporter = new OxyPlot.Wpf.PngExporter { Width = 1000, Height = 500 };
                using var fs = File.Create(dlg.FileName);
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
        public ObservableCollection<(int Cycle, double Value)> Points { get; set; } = new();
    }
}
