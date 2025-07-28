using ClosedXML.Excel;
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
        public IRelayCommand ExportCommand { get; }
        public IRelayCommand OpenTemplate { get; }
        [ObservableProperty]
        private string _exceedTotalSpDisplay = string.Empty;

        [ObservableProperty]
        private string _exceedTotalCalcDisplay = string.Empty;
        [ObservableProperty]
        private string _exceedRelSpDisplay = string.Empty;


        [ObservableProperty]
        private string _exceedRelCalcDisplay = string.Empty;


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
            ExportCommand = new RelayCommand(DoExport, () => Report is not null);

            _raw.PropertyChanged += (_, e) =>
            {
                if (e.PropertyName is nameof(RawDataViewModel.DataRows)
                                  or nameof(RawDataViewModel.Header.MaxNomen)
                                  or nameof(RawDataViewModel.Header.MaxCalculated)
                                  or nameof(RawDataViewModel.Header.RelNomen)
                                  or nameof(RawDataViewModel.Header.RelCalculated))
                {
                    Recalc();
                    ExportCommand.NotifyCanExecuteChanged();
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
                            ? $"{id}({row.Total:F2})"
                            : id;
                    })
                );

                ExceedTotalCalcDisplay = string.Join(", ",
                    Report.ExceedTotalCalcIds.Select(id =>
                    {
                        var row = _raw.DataRows.FirstOrDefault(r => r.Id == id);
                        return row != null
                            ? $"{id}({row.Total:F2})"
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
                              .Select(r => $"{r.Id1}-{r.Id2}")
                              .ToList();
            ExceedRelCalc = rel.ExceededCalcRows
                              .Select(r => $"{r.Id1}-{r.Id2}")
                              .ToList(); ExceedRelSpDisplay = string.Join(", ",
                rel.ExceededSpRows
                   .Select(r => $"{r.Id1}-{r.Id2}({r.Ratio:F4})")
            );

            ExceedRelCalcDisplay = string.Join(", ",
                rel.ExceededCalcRows
                   .Select(r => $"{r.Id1}-{r.Id2}({r.Ratio:F4})")
            );
            ExportCommand.NotifyCanExecuteChanged();
        }



        private void DoExport()
        {
            if (Report is null) return;

            var dlg = new SaveFileDialog
            {
                Filter = "Excel|*.xlsx",
                FileName = $"General_{DateTime.Now:yyyyMMdd_HHmm}.xlsx"
            };
            if (dlg.ShowDialog() != true) return;

            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("General");

            ws.Cell(1, 1).Value = "MAX Total"; ws.Cell(1, 2).Value = Report.MaxTotal.Value;
            ws.Cell(2, 1).Value = "Id(s)"; ws.Cell(2, 2).Value = string.Join(", ", Report.MaxTotal.Ids);

            ws.Cell(4, 1).Value = "MIN Total"; ws.Cell(4, 2).Value = Report.MinTotal.Value;
            ws.Cell(5, 1).Value = "Id(s)"; ws.Cell(5, 2).Value = string.Join(", ", Report.MinTotal.Ids);

            ws.Cell(7, 1).Value = "AVG Total"; ws.Cell(7, 2).Value = Report.AvgTotal;

            ws.Cell(9, 1).Value = "MAX Settl"; ws.Cell(9, 2).Value = Report.MaxSettl.Value;
            ws.Cell(10, 1).Value = "Id(s)"; ws.Cell(10, 2).Value = string.Join(", ", Report.MaxSettl.Ids);

            ws.Cell(12, 1).Value = "MIN Settl"; ws.Cell(12, 2).Value = Report.MinSettl.Value;
            ws.Cell(13, 1).Value = "Id(s)"; ws.Cell(13, 2).Value = string.Join(", ", Report.MinSettl.Ids);

            ws.Cell(15, 1).Value = "AVG Settl"; ws.Cell(15, 2).Value = Report.AvgSettl;

            ws.Cell(17, 1).Value = "Нет доступа:"; ws.Cell(17, 2).Value = string.Join(", ", Report.NoAccessIds);
            ws.Cell(18, 1).Value = "Уничтожены:"; ws.Cell(18, 2).Value = string.Join(", ", Report.DestroyedIds);
            ws.Cell(19, 1).Value = "Новые:"; ws.Cell(19, 2).Value = string.Join(", ", Report.NewIds);

            var spList = Report.ExceedTotalSpIds.Select(id =>
            {
                var row = _raw.DataRows.FirstOrDefault(r => r.Id == id);
                return row != null
                    ? $"{id}({row.Total:F2})"
                    : id;
            });
            ws.Cell(21, 1).Value = "Total > СП:";
            ws.Cell(21, 2).Value = string.Join(", ", spList); ws.Cell(22, 1).Value = "Total > расчёт:"; ws.Cell(22, 2).Value = string.Join(", ", Report.ExceedTotalCalcIds);

            var calcList = Report.ExceedTotalCalcIds.Select(id =>
            {
                var row = _raw.DataRows.FirstOrDefault(r => r.Id == id);
                return row != null
                    ? $"{id}({row.Total:F2})"
                    : id;
            });
            ws.Cell(22, 1).Value = "Total > расчёт:";
            ws.Cell(22, 2).Value = string.Join(", ", calcList);

            ws.Cell(24, 1).Value = "Превышения относительной осадки по СП:";
            ws.Cell(24, 2).Value = string.Join(", ", ExceedRelSp);

            ws.Cell(25, 1).Value = "Превышения относительной осадки по расчётам:";
            ws.Cell(25, 2).Value = string.Join(", ", ExceedRelCalc);


            wb.SaveAs(dlg.FileName);
        }
    }
}
