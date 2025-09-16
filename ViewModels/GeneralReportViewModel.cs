﻿using ClosedXML.Excel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Microsoft.Win32;
using Osadka.Models;
using Osadka.Services;
using System;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
using System.Windows;
using System.IO;

namespace Osadka.ViewModels
{
    public partial class GeneralReportViewModel : ObservableObject
    {
        public IRelayCommand CalculateCommand { get; }
        public IRelayCommand OpenTemplate { get; }
        [ObservableProperty]
        private string _exceedTotalSpDisplay = string.Empty;

        [ObservableProperty]
        private string _exceedTotalCalcDisplay = string.Empty;
        [ObservableProperty]
        private string _exceedRelSpDisplay = string.Empty;

        [ObservableProperty]
        private string _exceedRelCalcDisplay = string.Empty;
        public ReportOutputSettings Settings { get; } = new();
        [ObservableProperty]
        private GeneralReportData? _report;

        [ObservableProperty]
        private IReadOnlyList<string> _exceedRelSp = Array.Empty<string>();

        [ObservableProperty]
        private IReadOnlyList<string> _exceedRelCalc = Array.Empty<string>();

        private readonly RawDataViewModel _raw;
        private readonly GeneralReportService _svc;
        private readonly RelativeReportService _relSvc;

        public GeneralReportViewModel(
            RawDataViewModel raw,
            GeneralReportService svc,
            RelativeReportService relSvc)
        {
            _raw = raw ?? throw new ArgumentNullException(nameof(raw));
            _svc = svc ?? throw new ArgumentNullException(nameof(svc));
            _relSvc = relSvc ?? throw new ArgumentNullException(nameof(relSvc));

            CalculateCommand = new RelayCommand(Recalc);
            OpenTemplate = new RelayCommand(Opentemp);


            _raw.PropertyChanged += (_, e) =>
            {
                if (e.PropertyName is nameof(RawDataViewModel.DataRows)
                                  or nameof(RawDataViewModel.Header.MaxNomen)
                                  or nameof(RawDataViewModel.Header.MaxCalculated)
                                  or nameof(RawDataViewModel.Header.RelNomen)
                                  or nameof(RawDataViewModel.Header.RelCalculated))
                {
                    Recalc();

                }
            };
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
            _raw.DataRows.CollectionChanged += (_, __) => Recalc();
            _raw.CoordRows.CollectionChanged += (_, __) => Recalc();

            Recalc();
        }
        private void Opentemp()
        {
            string exeDir = AppContext.BaseDirectory;
            string template = Path.Combine(exeDir, "template.xlsx");
            if (!File.Exists(template))
            {
                MessageBox.Show("template.xlsx не найден", "Экспорт",
                                MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            if (File.Exists(template))
                Process.Start(new ProcessStartInfo(template) { UseShellExecute = true });
            else
                MessageBox.Show("Файл справки не найден.",
                                "Справка",
                                MessageBoxButton.OK,
                                MessageBoxImage.Warning);
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
                    Report.ExceedTotalSpIds.Select(id =>
                    {
                        var row = _raw.DataRows.FirstOrDefault(r => r.Id == id);
                        return row != null
                            ? $"№{id}({row.Total:F1})"
                            : id;
                    })
                );

                ExceedTotalCalcDisplay = string.Join(", ",
                    Report.ExceedTotalCalcIds.Select(id =>
                    {
                        var row = _raw.DataRows.FirstOrDefault(r => r.Id == id);
                        return row != null
                            ? $"№{id}({row.Total:F1})"
                            : id;
                    })
                );
            }
            else
            {
                ExceedTotalSpDisplay = string.Empty;
                ExceedTotalCalcDisplay = string.Empty;
            }

            double spLim = _raw.Header.RelNomen ?? 0;
            double calcLim = _raw.Header.RelCalculated ?? 0;
            var rel = _relSvc.Build(
                _raw.CoordRows,
                _raw.DataRows,
                spLim,
                calcLim);
            ExceedRelSp = rel.ExceededSpRows
                              .Select(r => $"№ {r.Id1}-{r.Id2}")
                              .ToList();
            ExceedRelCalc = rel.ExceededCalcRows
                              .Select(r => $"№{r.Id1}-{r.Id2}")
                              .ToList(); ExceedRelSpDisplay = string.Join(", ",
                rel.ExceededSpRows
                   .Select(r => $"{r.Id1}-{r.Id2}({r.Ratio:F5})")
            );

            ExceedRelCalcDisplay = string.Join(", ",
                rel.ExceededCalcRows
                   .Select(r => $"{r.Id1}-{r.Id2}({r.Ratio:F5})")
            );
        }

    }
}
