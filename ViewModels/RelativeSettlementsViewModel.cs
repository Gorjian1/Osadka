using System;
using System.Linq;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using ClosedXML.Excel;
using Microsoft.Win32;
using Osadka.Models;
using Osadka.Services;

namespace Osadka.ViewModels
{
    public partial class RelativeSettlementsViewModel : ObservableObject
    {
        public ObservableCollection<RelativeRow> AllRows { get; } = new();
        public ObservableCollection<RelativeRow> ExceededSpRows { get; } = new();
        public ObservableCollection<RelativeRow> ExceededCalcRows { get; } = new();

        public IRelayCommand RecalcCommand { get; }
        public IRelayCommand ExportCommand { get; }

        private readonly RawDataViewModel _raw;
        private readonly RelativeReportService _relSvc;

        public RelativeSettlementsViewModel(RawDataViewModel raw, RelativeReportService svc)
        {
            _raw = raw ?? throw new ArgumentNullException(nameof(raw));
            _relSvc = svc ?? throw new ArgumentNullException(nameof(svc));

            RecalcCommand = new RelayCommand(Recalc);
            ExportCommand = new RelayCommand(DoExport, () => AllRows.Any());

            _raw.PropertyChanged += (_, e) =>
            {
                if (e.PropertyName is nameof(RawDataViewModel.Header.ObjectNumber)
                                     or nameof(RawDataViewModel.DataRows)
                                     or nameof(RawDataViewModel.CoordRows))
                    Recalc();
            };
            _raw.DataRows.CollectionChanged += (_, __) => Recalc();
            _raw.CoordRows.CollectionChanged += (_, __) => Recalc();

            Recalc();
        }

        private void Recalc()
        {
            double spLim = _raw.Header.RelNomen ?? 0;
            double calcLim = _raw.Header.RelCalculated ?? 0;

            var report = _relSvc.Build(
                _raw.CoordRows,
                _raw.DataRows,
                spLim,
                calcLim);

            AllRows.Reset(report.AllRows);
            ExceededSpRows.Reset(report.ExceededSpRows);
            ExceededCalcRows.Reset(report.ExceededCalcRows);

            ExportCommand.NotifyCanExecuteChanged();
        }

        private void DoExport()
        {
            var dlg = new SaveFileDialog
            {
                Filter = "Excel|*.xlsx",
                FileName = $"Relative_{DateTime.Now:yyyyMMdd_HHmm}.xlsx"
            };
            if (dlg.ShowDialog() != true) return;

            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Relative");

            ws.Cell(1, 1).Value = "Точка №1";
            ws.Cell(1, 2).Value = "Точка №1";
            ws.Cell(1, 3).Value = "Расстояние, мм";
            ws.Cell(1, 4).Value = "Абс. Разность, мм";
            ws.Cell(1, 5).Value = "Отн. Разность";

            int row = 2;
            foreach (var r in AllRows)
            {
                ws.Cell(row, 1).Value = r.Id1;
                ws.Cell(row, 2).Value = r.Id2;
                ws.Cell(row, 3).Value = r.Distance;
                ws.Cell(row, 4).Value = r.DeltaTotal;
                ws.Cell(row, 5).Value = r.Ratio;
                row++;
            }

            row++;
            ws.Cell(row, 1).Value = $"Превышения по СП > {(_raw.Header.RelNomen ?? 0)})";
            row++;
            foreach (var r in ExceededSpRows)
            {
                ws.Cell(row, 1).Value = $"{r.Id1}-{r.Id2}";
                row++;
            }

            row++;
            ws.Cell(row, 1).Value = $"Превышения по Расчетам {(_raw.Header.RelCalculated ?? 0)})";
            row++;
            foreach (var r in ExceededCalcRows)
            {
                ws.Cell(row, 1).Value = $"{r.Id1}-{r.Id2}";
                row++;
            }
            int lastDataRow = row - 1;
            var tableRange = ws.Range(1, 1, lastDataRow, 5);

            tableRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            tableRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;

            tableRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            ws.Row(1).Style.Font.Bold = true;
            ws.Columns(1, 5).AdjustToContents();
            wb.SaveAs(dlg.FileName);
        }
    }

    static class Extensions
    {
        public static void Reset<T>(this ObservableCollection<T> coll, IEnumerable<T> src)
        {
            coll.Clear();
            foreach (var i in src) coll.Add(i);
        }
    }
}
