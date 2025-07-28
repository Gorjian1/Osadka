using ClosedXML.Excel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using DocumentFormat.OpenXml.Drawing.Charts;
using Microsoft.Win32;
using Osadka.Models;
using Osadka.Services;
using Osadka.Views;
using OxyPlot;
using OxyPlot.Wpf;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net.NetworkInformation;
using System.Reflection;
using System.Text.Json;
using System.Windows;


namespace Osadka.ViewModels
{
    public partial class MainViewModel : ObservableObject
    {
        private string? _currentPath;

        public RawDataViewModel RawVM { get; }
        public GeneralReportViewModel GenVM { get; }
        public RelativeSettlementsViewModel RelVM { get; }
        public DynamicsGrafficViewModel DynVM { get; }
        private readonly DynamicsReportService _dynSvc;
        public IRelayCommand HelpCommand { get; }
        public IRelayCommand<string> NavigateCommand { get; }
        public IRelayCommand NewProjectCommand { get; }
        public IRelayCommand OpenProjectCommand { get; }
        public IRelayCommand SaveProjectCommand { get; }
        public IRelayCommand SaveAsProjectCommand { get; }
        public IRelayCommand QuickReportCommand { get; }
        private CoordinateExporting? _coord;

        private object? _currentPage;
        public object? CurrentPage
        {
            get => _currentPage;
            set => SetProperty(ref _currentPage, value);
        }
        private static class PageKeys
        {
            public const string Raw = "RawData";
            public const string Diff = "SettlementDiff";
            public const string Sum = "Summary";
            public const string Graf = "Graffics";
            public const string Coord = "Coordinates";
        }
        private void OpenHelp()
        {
            string exeDir = AppContext.BaseDirectory;
            string docx = Path.Combine(exeDir, "help.docx");

            if (File.Exists(docx))
                Process.Start(new ProcessStartInfo(docx) { UseShellExecute = true });
            else
                MessageBox.Show("Файл справки не найден.",
                                "Справка",
                                MessageBoxButton.OK,
                                MessageBoxImage.Warning);
        }
        public MainViewModel()
        {
            RawVM = new RawDataViewModel();

            var genSvc = new GeneralReportService();
            var relSvc = new RelativeReportService();

            GenVM = new GeneralReportViewModel(RawVM, genSvc, relSvc);
            RelVM = new RelativeSettlementsViewModel(RawVM, relSvc);
            _dynSvc = new DynamicsReportService();

            HelpCommand = new RelayCommand(OpenHelp);
            NavigateCommand = new RelayCommand<string>(Navigate);
            NewProjectCommand = new RelayCommand(NewProject);
            OpenProjectCommand = new RelayCommand(OpenProject);
            SaveProjectCommand = new RelayCommand(SaveProject);
            SaveAsProjectCommand = new RelayCommand(SaveAsProject);
            QuickReportCommand = new RelayCommand(DoQuickExport, () => GenVM.Report != null);

            Navigate(PageKeys.Raw);
        }

        #region Navigation

        private void Navigate(string? key)
        {
            CurrentPage = key switch
            {
                PageKeys.Raw => new RawDataPage(RawVM),
                PageKeys.Diff => new GeneralReportPage(GenVM),
                PageKeys.Sum => new RelativeSettlementsPage(RelVM),
                PageKeys.Coord => _coord ??= new CoordinateExporting(RawVM),
                PageKeys.Graf => new DynamicsGrafficPage(new DynamicsGrafficViewModel(RawVM, _dynSvc)),
                _ => CurrentPage
            };
        }

        #endregion


        private void NewProject()
        {
            RawVM.ClearCommand.Execute(null);
            _currentPath = null;
        }

        private void OpenProject()
        {
            var dlg = new OpenFileDialog
            {
                Filter = "Osadka Project (*.data)|*.data|All Files|*.*"
            };
            if (dlg.ShowDialog() != true) return;
            if (RawVM is not { } vm) return;

            try
            {
                var json = File.ReadAllText(dlg.FileName);
                var data = JsonSerializer.Deserialize<ProjectData>(json)
                           ?? throw new InvalidOperationException("Невалидный формат");


                vm.Header.CycleNumber = data.Cycle;
                vm.Header.MaxNomen = data.MaxNomen;
                vm.Header.MaxCalculated = data.MaxCalculated;
                vm.Header.RelNomen = data.RelNomen;
                vm.Header.RelCalculated = data.RelCalculated;

                vm.DataRows.Clear();
                foreach (var r in data.DataRows) vm.DataRows.Add(r);
                vm.CoordRows.Clear();
                foreach (var c in data.CoordRows) vm.CoordRows.Add(c);


                vm.Objects.Clear();
                foreach (var objKv in data.Objects)
                {
                    vm.Objects[objKv.Key] = objKv.Value.ToDictionary(
                        cycleKv => cycleKv.Key,
                        cycleKv => cycleKv.Value);
                }
                vm.ObjectNumbers.Clear();
                foreach (var obj in vm.Objects.Keys.OrderBy(k => k))
                    vm.ObjectNumbers.Add(obj);

                vm.CycleNumbers.Clear();
                if (vm.Objects.TryGetValue(vm.Header.ObjectNumber, out var cycles))
                    foreach (var cyc in cycles.Keys.OrderBy(k => k))
                        vm.CycleNumbers.Add(cyc);

                _currentPath = dlg.FileName;
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show(
                    $"Ошибка при загрузке проекта:\n{ex.Message}",
                    "Ошибка",
                    System.Windows.MessageBoxButton.OK,
                    System.Windows.MessageBoxImage.Error);
            }
        }

        private void SaveProject()
        {
            if (_currentPath == null)
            {
                SaveAsProject();
                return;
            }
            SaveTo(_currentPath);
        }

        private void SaveAsProject()
        {
            var dlg = new SaveFileDialog
            {
                Filter = "Osadka Project (*.data)|*.data"
            };
            if (dlg.ShowDialog() != true) return;
            SaveTo(dlg.FileName);
            _currentPath = dlg.FileName;
        }

        private void SaveTo(string path)
        {
            if (RawVM is not { } vm) return;

            var data = new ProjectData
            {
                Cycle = vm.Header.CycleNumber,
                MaxNomen = vm.Header.MaxNomen,
                MaxCalculated = vm.Header.MaxCalculated,
                RelNomen = vm.Header.RelNomen,
                RelCalculated = vm.Header.RelCalculated,

                DataRows = vm.DataRows.ToList(),
                CoordRows = vm.CoordRows.ToList(),

                Objects = vm.Objects.ToDictionary(
                    objKv => objKv.Key,
                    objKv => objKv.Value.ToDictionary(
                        cycleKv => cycleKv.Key,
                        cycleKv => cycleKv.Value.ToList()
                    ))
            };

            var json = JsonSerializer.Serialize(
                data,
                new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(path, json);
        }


        private void DoQuickExport()
        {
            if (GenVM.Report is null) return;

            string exeDir = AppContext.BaseDirectory;
            string template = Path.Combine(exeDir, "template.xlsx");
            if (!File.Exists(template))
            {
                MessageBox.Show("template.xlsx не найден", "Экспорт",
                                MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            var dlg = new SaveFileDialog
            {
                Filter = "Excel|*.xlsx",
                FileName = $"БОтчёт_{DateTime.Now:yyyyMMdd_HHmm}.xlsx"
            };
            if (dlg.ShowDialog() != true) return;

            File.Copy(template, dlg.FileName, overwrite: true);

            using var wb = new XLWorkbook(dlg.FileName);

            var map = BuildPlaceholderMap(); 
            foreach (var cell in wb.Worksheets.First()
                           .CellsUsed(c => c.DataType == XLDataType.Text))
            {
                string t = cell.GetString().Trim();
                if (t.StartsWith("/") && map.TryGetValue(t, out var val))
                    cell.Value = val;
            }

            AddRelativeSheet(wb);
            AddDynamicsSheet(wb);

            wb.Save();
        }

        private Dictionary<string, string> BuildPlaceholderMap()
        {
            var map = new Dictionary<string, string>
            {
                ["/цикл"] = RawVM.SelectedCycleHeader,

                ["/предСПмакс"] = RawVM.Header.MaxNomen?.ToString("F2") ?? "-",
                ["/предРАСЧмакс"] = RawVM.Header.MaxCalculated?.ToString("F2") ?? "-",
                ["/предСПотн"] = RawVM.Header.RelNomen?.ToString("F2") ?? "-",
                ["/предРАСЧотн"] = RawVM.Header.RelCalculated?.ToString("F2") ?? "-",

                ["/общмакс"] = $"{GenVM.Report.MaxTotal.Value:F2}",
                ["/общмаксId"] = string.Join(", ", GenVM.Report.MaxTotal.Ids),
                ["/общмин"] = $"{GenVM.Report.MinTotal.Value:F2}",
                ["/общминId"] = string.Join(", ", GenVM.Report.MinTotal.Ids),
                ["/общср"] = $"{GenVM.Report.AvgTotal:F2}",

                ["/сеттмакс"] = $"{GenVM.Report.MaxSettl.Value:F2}",
                ["/сеттмаксId"] = string.Join(", ", GenVM.Report.MaxSettl.Ids),
                ["/сеттмин"] = $"{GenVM.Report.MinSettl.Value:F2}",
                ["/сеттминId"] = string.Join(", ", GenVM.Report.MinSettl.Ids),
                ["/сеттср"] = $"{GenVM.Report.AvgSettl:F2}",

                ["/нетдоступа"] = string.Join(", ", GenVM.Report.NoAccessIds),
                ["/уничтожены"] = string.Join(", ", GenVM.Report.DestroyedIds),
                ["/новые"] = string.Join(", ", GenVM.Report.NewIds),

                ["/общ>сп"] = GenVM.ExceedTotalSpDisplay,
                ["/общ>расч"] = GenVM.ExceedTotalCalcDisplay,
                ["/отн>сп"] = GenVM.ExceedRelSpDisplay,
                ["/отн>расч"] = GenVM.ExceedRelCalcDisplay
            };
            return map;
        }

        private void AddRelativeSheet(XLWorkbook wb)
        {
            var ws = wb.AddWorksheet("Относительная разность");
            ws.Cell(1, 1).Value = "Точка №1";
            ws.Cell(1, 2).Value = "Точка №2";
            ws.Cell(1, 3).Value = "Расстояние, мм";
            ws.Cell(1, 4).Value = "Абс. Разность, мм";
            ws.Cell(1, 5).Value = "Отн. Разность";
            int r = 2;
            foreach (var row in RelVM.AllRows)
            {
                ws.Cell(r, 1).Value = row.Id1;
                ws.Cell(r, 2).Value = row.Id2;
                ws.Cell(r, 3).Value = row.Distance;
                ws.Cell(r, 4).Value = row.DeltaTotal;
                ws.Cell(r, 5).Value = row.Ratio;
                r++;
            }


            int lastDataRow = r - 1;
            var rng = ws.Range(1, 1, lastDataRow, 5);

            rng.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            rng.Style.Border.InsideBorder = XLBorderStyleValues.Thin;

            rng.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            ws.Row(1).Style.Font.Bold = true;
            ws.Columns(1, 5).AdjustToContents();

        }

        private void AddDynamicsSheet(XLWorkbook wb)
        {
            var dynVm = new DynamicsGrafficViewModel(RawVM, _dynSvc);
            var ws = wb.AddWorksheet("Графики динамики");
            ws.Cell(1, 1).Value = "Id";
            var cycles = RawVM.CurrentCycles.Keys.OrderBy(c => c).ToList();
            for (int i = 0; i < cycles.Count; i++)
                ws.Cell(1, i + 2).Value = $"Cycle {cycles[i]}";

            int r = 2;
            foreach (var ser in dynVm.Lines)
            {
                ws.Cell(r, 1).Value = ser.Id;
                foreach (var pt in ser.Points)
                {
                    int c = cycles.IndexOf(pt.Cycle) + 2;
                    ws.Cell(r, c).Value = pt.Mark;
                }
                r++;
            }

            using var ms = new MemoryStream();

            var tmpFile = Path.GetTempFileName();
            var png = new OxyPlot.Wpf.PngExporter
            {
                Width = 800,
                Height = 400,
            }
            ;
             using (var fs = File.OpenWrite(tmpFile))
                png.Export(dynVm.PlotModel, fs);

            ws.AddPicture(tmpFile)
              .MoveTo(ws.Cell(1, cycles.Count + 4))
              .WithSize(800, 400);

            File.Delete(tmpFile);
        }


    }
}
