using ClosedXML.Excel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Microsoft.Win32;
using Osadka.Models;
using Osadka.Services;
using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Windows;

namespace Osadka.ViewModels
{
    public partial class GeneralReportViewModel : ObservableObject
    {
        public IRelayCommand CalculateCommand { get; }
        public IRelayCommand OpenTemplate { get; }

        [ObservableProperty] private string _exceedTotalSpDisplay = string.Empty;
        [ObservableProperty] private string _exceedTotalCalcDisplay = string.Empty;
        [ObservableProperty] private string _exceedRelSpDisplay = string.Empty;
        [ObservableProperty] private string _exceedRelCalcDisplay = string.Empty;

        [ObservableProperty] private GeneralReportData? _report;
        [ObservableProperty] private IReadOnlyList<string> _exceedRelSp = Array.Empty<string>();
        [ObservableProperty] private IReadOnlyList<string> _exceedRelCalc = Array.Empty<string>();

        private readonly RawDataViewModel _raw;
        private readonly GeneralReportService _svc;


        public GeneralReportViewModel(RawDataViewModel raw, GeneralReportService svc)
        {
            _raw = raw ?? throw new ArgumentNullException(nameof(raw));
            _svc = svc ?? throw new ArgumentNullException(nameof(svc));

            CalculateCommand = new RelayCommand(Recalc);
            OpenTemplate = new RelayCommand(OpenTemplateFile);

            _raw.PropertyChanged += (_, e) =>
            {
                if (e.PropertyName is nameof(RawDataViewModel.DataRows)
                    or nameof(RawDataViewModel.CoordRows)
                    or nameof(RawDataViewModel.Header.MaxNomen)
                    or nameof(RawDataViewModel.Header.MaxCalculated)
                    or nameof(RawDataViewModel.Header.RelNomen)
                    or nameof(RawDataViewModel.Header.RelCalculated))
                {
                    Recalc();
                }
            };
            _raw.DataRows.CollectionChanged += (_, __) => Recalc();
            _raw.CoordRows.CollectionChanged += (_, __) => Recalc();
            raw.Header.PropertyChanged += (_, e) =>
            {
                if (e.PropertyName is nameof(CycleHeader.MaxNomen)
                    or nameof(CycleHeader.MaxCalculated)
                    or nameof(CycleHeader.RelNomen)
                    or nameof(CycleHeader.RelCalculated))
                {
                    Recalc();
                }
            };

            Recalc();
        }

        private void OpenTemplateFile()
        {
            string exeDir = AppContext.BaseDirectory;
            string template = Path.Combine(exeDir, "template.xlsx");
            if (!File.Exists(template))
            {
                MessageBox.Show("template.xlsx не найден", "Экспорт",
                                MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            Process.Start(new ProcessStartInfo(template) { UseShellExecute = true });
        }

        private void Recalc()
        {
            Report = _svc.Build(
                _raw.DataRows,
                _raw.Header.MaxNomen ?? 0,
                _raw.Header.MaxCalculated ?? 0);

            if (Report != null)
            {
                ExceedTotalSpDisplay = string.Join(", ",
                    Report.ExceedVectorSpIds.Select(id =>
                    {
                        var row = _raw.DataRows.FirstOrDefault(r => r.Id == id);
                        var val = row?.Vector ?? row?.Total;
                        return val.HasValue ? $"{id}({val.Value:F1})" : id;
                    }));

                ExceedTotalCalcDisplay = string.Join(", ",
                    Report.ExceedVectorCalcIds.Select(id =>
                    {
                        var row = _raw.DataRows.FirstOrDefault(r => r.Id == id);
                        var val = row?.Vector ?? row?.Total;
                        return val.HasValue ? $"{id}({val.Value:F1})" : id;
                    }));
            }
            else
            {
                ExceedTotalSpDisplay = string.Empty;
                ExceedTotalCalcDisplay = string.Empty;
            }

            double spLim = _raw.Header.RelNomen ?? 0;
            double calcLim = _raw.Header.RelCalculated ?? 0;


        }
    }
}
