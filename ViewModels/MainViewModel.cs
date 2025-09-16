// File: ViewModels/MainViewModel.cs
using ClosedXML.Excel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Microsoft.Win32;
using Osadka.Models;
using Osadka.Services;
using Osadka.Views;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Windows;
using Telegram.Bot.Types;

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
        private bool _includeGeneral = true;
        public bool IncludeGeneral
        {
            get => _includeGeneral;
            set => SetProperty(ref _includeGeneral, value);
        }

        private bool _includeRelative = true;
        public bool IncludeRelative
        {
            get => _includeRelative;
            set => SetProperty(ref _includeRelative, value);
        }

        private bool _includeGraphs = true;
        public bool IncludeGraphs
        {
            get => _includeGraphs;
            set => SetProperty(ref _includeGraphs, value);
        }

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

        public void LoadProject(string path)
        {
            if (RawVM is not { } vm) return;

            try
            {
                var json = File.ReadAllText(path);
                var data = JsonSerializer.Deserialize<ProjectData>(json)
                           ?? throw new InvalidOperationException("Невалидный формат");

                vm.Header.CycleNumber = data.Cycle;
                vm.Header.MaxNomen = data.MaxNomen;
                vm.Header.MaxCalculated = data.MaxCalculated;
                vm.Header.RelNomen = data.RelNomen;
                vm.Header.RelCalculated = data.RelCalculated;
                vm.SelectedCycleHeader = data.SelectedCycleHeader ?? string.Empty;

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

                _currentPath = path;
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"Ошибка при загрузке проекта:\n{ex.Message}",
                    "Ошибка",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
        }

        private void OpenProject()
        {
            var dlg = new OpenFileDialog
            {
                Filter = "Osadka Project (*.osd)|*.osd|All Files|*.*"
            };
            if (dlg.ShowDialog() != true) return;

            LoadProject(dlg.FileName);
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
                Filter = "Osadka Project (*.osd)|*.osd"
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
                SelectedCycleHeader = vm.SelectedCycleHeader,
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

            if (!(IncludeGeneral || IncludeRelative || IncludeGraphs))
            {
                MessageBox.Show("Выберите хотя бы один пункт: Общий, Относительный, Графики.",
                                "Экспорт", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            // Выбор шаблона: сначала пользовательский, затем встроенный
            string? userTemplate = RawVM?.TemplatePath;
            string exeDir = AppContext.BaseDirectory;
            string fallbackTemplate = Path.Combine(exeDir, "template.xlsx");
            string template = (!string.IsNullOrWhiteSpace(userTemplate) && File.Exists(userTemplate))
                 ? userTemplate!
                 : fallbackTemplate;

            if (!File.Exists(template))
            {
                MessageBox.Show("Шаблон Excel не найден: ни выбранный, ни встроенный template.xlsx.",
                "Экспорт", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            // Подстраиваем расширение под шаблон (.xlsx или .xlsm)
            string ext = Path.GetExtension(template).Equals(".xlsm", StringComparison.OrdinalIgnoreCase) ? "*.xlsm" : "*.xlsx";
            string filter = $"Excel ({ext})|{ext}";
            var dlg = new SaveFileDialog
            {
                Filter = filter,
                FileName = $"Отчёт_{DateTime.Now:yyyyMMdd_HHmm}{Path.GetExtension(template)}"
            };

            if (dlg.ShowDialog() != true) return;

            try
            {
                File.Copy(template, dlg.FileName, overwrite: true);

                using (var wb = new XLWorkbook(dlg.FileName))
                {
                    var generalWs = wb.Worksheets.First(); // титульный/общий

                    if (IncludeGeneral)
                    {
                        var map = BuildPlaceholderMap();

                        // ВАЖНО: поддержка выключенных блоков — удаляем строки, где встретились их теги
                        var disabled = GenVM.Settings?.GetDisabledTags()
                                       ?? new System.Collections.Generic.HashSet<string>();

                        // Снимок текстовых ячеек, чтобы удалять строки уже после прохода
                        var textCells = generalWs.CellsUsed(c => c.DataType == XLDataType.Text).ToList();
                        var rowsToDelete = new System.Collections.Generic.HashSet<int>();

                        foreach (var cell in textCells)
                        {
                            string t = cell.GetString().Trim();
                            if (!t.StartsWith("/")) continue;

                            if (disabled.Contains(t))
                            {
                                rowsToDelete.Add(cell.Address.RowNumber);
                                continue;
                            }

                            if (map.TryGetValue(t, out var val))
                                cell.Value = val;
                        }

                        // Удаляем строки снизу вверх
                        foreach (var r in rowsToDelete.OrderByDescending(x => x))
                            generalWs.Row(r).Delete();
                    }
                    else
                    {
                        generalWs.Delete();
                    }

                    if (IncludeRelative)
                        AddRelativeSheet(wb);

                    if (IncludeGraphs)
                    {
                        AddDynamicsSheet(wb);
                    }
                    else
                    {
                        var dynWs = wb.Worksheets.FirstOrDefault(
                            ws => string.Equals(ws.Name, "Графики динамики", System.StringComparison.OrdinalIgnoreCase));
                        dynWs?.Delete();
                    }

                    wb.Save();
                }

                if (IncludeGraphs)
                {
                    RunSta(() => BuildChartFromDynTable_Quick_NoPIA(
                        filePath: dlg.FileName,
                        dataSheetName: "Графики динамики",
                        tableName: "DynTable",
                        chartSheetName: "Графики динамики",
                        left: 40, top: 200, width: 920, height: 440,
                        deleteOldCharts: true
                    ));
                }

                MessageBox.Show("Экспорт завершён", "Экспорт",
                                MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка экспорта: {ex.Message}", "Экспорт",
                                MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private static void RunSta(System.Action action)
        {
            var t = new System.Threading.Thread(() => action()) { IsBackground = true };
            t.SetApartmentState(System.Threading.ApartmentState.STA);
            t.Start();
            t.Join();
        }

        private static void BuildChartFromDynTable_Quick_NoPIA(
            string filePath,
            string dataSheetName = "Графики динамики",
            string tableName = "DynTable",
            string chartSheetName = "Графики динамики",
            int left = 40, int top = 200, int width = 920, int height = 440,
            bool deleteOldCharts = true)
        {
            const int xlRows = 1;
            const int xlLine = 4;

            object app = null, wb = null, wsData = null, wsChart = null;
            object workbooks = null, worksheets = null;
            object listObjects = null, lo = null, loRange = null, dataBodyRange = null;
            object chartObjects = null, chartObj = null, chart = null;

            try
            {
                var excelType = System.Type.GetTypeFromProgID("Excel.Application", throwOnError: false);
                if (excelType == null)
                    throw new System.InvalidOperationException("На этом ПК не установлен Microsoft Excel.");

                app = System.Activator.CreateInstance(excelType);
                excelType.InvokeMember("Visible",
                    System.Reflection.BindingFlags.SetProperty, null, app, new object[] { false });
                excelType.InvokeMember("DisplayAlerts",
                    System.Reflection.BindingFlags.SetProperty, null, app, new object[] { false });

                // Workbooks.Open(filePath)
                workbooks = excelType.InvokeMember("Workbooks",
                    System.Reflection.BindingFlags.GetProperty, null, app, null);
                wb = workbooks.GetType().InvokeMember("Open",
                    System.Reflection.BindingFlags.InvokeMethod, null, workbooks, new object[] { filePath });

                // Worksheets[dataSheetName]
                worksheets = wb.GetType().InvokeMember("Worksheets",
                    System.Reflection.BindingFlags.GetProperty, null, wb, null);
                try
                {
                    wsData = worksheets.GetType().InvokeMember("Item",
                        System.Reflection.BindingFlags.GetProperty, null, worksheets, new object[] { dataSheetName });
                }
                catch
                {
                    throw new System.InvalidOperationException($"Не найден лист данных '{dataSheetName}'.");
                }

                // Таблица ListObjects[tableName]
                listObjects = wsData.GetType().InvokeMember("ListObjects",
                    System.Reflection.BindingFlags.GetProperty, null, wsData, null);
                try
                {
                    lo = listObjects.GetType().InvokeMember("Item",
                        System.Reflection.BindingFlags.GetProperty, null, listObjects, new object[] { tableName });
                }
                catch
                {
                    throw new System.InvalidOperationException($"Не найдена таблица '{tableName}' на листе '{dataSheetName}'.");
                }

                dataBodyRange = lo.GetType().InvokeMember("DataBodyRange",
                    System.Reflection.BindingFlags.GetProperty, null, lo, null);
                if (dataBodyRange == null)
                    throw new System.InvalidOperationException("В таблице нет строк данных.");
                loRange = lo.GetType().InvokeMember("Range",
                    System.Reflection.BindingFlags.GetProperty, null, lo, null);

                try
                {
                    wsChart = worksheets.GetType().InvokeMember("Item",
                        System.Reflection.BindingFlags.GetProperty, null, worksheets, new object[] { chartSheetName });
                }
                catch
                {
                    wsChart = worksheets.GetType().InvokeMember("Add",
                        System.Reflection.BindingFlags.InvokeMethod, null, worksheets, null);
                    wsChart.GetType().InvokeMember("Name",
                        System.Reflection.BindingFlags.SetProperty, null, wsChart, new object[] { chartSheetName });
                }
                object missing = System.Type.Missing;
                try
                {
                    chartObjects = wsChart.GetType().InvokeMember("ChartObjects",
                        System.Reflection.BindingFlags.InvokeMethod, null, wsChart, new object[] { missing });
                }
                catch
                {
                    chartObjects = wsChart.GetType().InvokeMember("ChartObjects",
                        System.Reflection.BindingFlags.InvokeMethod, null, wsChart, null);
                }

                if (deleteOldCharts && chartObjects != null)
                {
                    try
                    {
                        chartObjects.GetType().InvokeMember("Delete",
                            System.Reflection.BindingFlags.InvokeMethod, null, chartObjects, null);
                    }
                    catch
                    {
                        try
                        {
                            var cntObj = chartObjects.GetType().InvokeMember("Count",
                                System.Reflection.BindingFlags.GetProperty, null, chartObjects, null);
                            int cnt = cntObj is int i ? i : System.Convert.ToInt32(cntObj);
                            for (int j = cnt; j >= 1; j--)
                            {
                                var co = chartObjects.GetType().InvokeMember("Item",
                                    System.Reflection.BindingFlags.GetProperty, null, chartObjects, new object[] { j });
                                co.GetType().InvokeMember("Delete",
                                    System.Reflection.BindingFlags.InvokeMethod, null, co, null);
                                Marshal.FinalReleaseComObject(co);
                            }
                        }
                        catch { }
                    }

                    try
                    {
                        chartObjects = wsChart.GetType().InvokeMember("ChartObjects",
                            System.Reflection.BindingFlags.InvokeMethod, null, wsChart, new object[] { missing });
                    }
                    catch
                    {
                        chartObjects = wsChart.GetType().InvokeMember("ChartObjects",
                            System.Reflection.BindingFlags.InvokeMethod, null, wsChart, null);
                    }
                }

                chartObj = chartObjects.GetType().InvokeMember("Add",
                    System.Reflection.BindingFlags.InvokeMethod, null, chartObjects,
                    new object[] { left, top, width, height });
                chart = chartObj.GetType().InvokeMember("Chart",
                    System.Reflection.BindingFlags.GetProperty, null, chartObj, null);

                chart.GetType().InvokeMember("SetSourceData",
                    System.Reflection.BindingFlags.InvokeMethod, null, chart, new object[] { loRange, xlRows });
                chart.GetType().InvokeMember("ChartType",
                    System.Reflection.BindingFlags.SetProperty, null, chart, new object[] { xlLine });
                chart.GetType().InvokeMember("HasTitle",
                    System.Reflection.BindingFlags.SetProperty, null, chart, new object[] { false });

                wb.GetType().InvokeMember("Save", System.Reflection.BindingFlags.InvokeMethod, null, wb, null);
            }
            finally
            {
                try { wb?.GetType().InvokeMember("Close", System.Reflection.BindingFlags.InvokeMethod, null, wb, new object[] { false }); } catch { }
                try { app?.GetType().InvokeMember("Quit", System.Reflection.BindingFlags.InvokeMethod, null, app, null); } catch { }

                void rel(object o) { if (o != null) Marshal.FinalReleaseComObject(o); }
                rel(chart); rel(chartObj); rel(chartObjects);
                rel(wsChart); rel(loRange); rel(dataBodyRange); rel(lo); rel(listObjects);
                rel(wsData); rel(worksheets); rel(wb); rel(workbooks); rel(app);

                System.GC.Collect(); System.GC.WaitForPendingFinalizers();
                System.GC.Collect(); System.GC.WaitForPendingFinalizers();
            }
        }

        private Dictionary<string, string> BuildPlaceholderMap()
        {
            static string DashIfEmpty(string? s) =>
                string.IsNullOrWhiteSpace(s) ? "-" : s;

            static string JoinOrDash(IEnumerable<string>? ids)
            {
                if (ids == null) return "-";
                var arr = ids.Where(s => !string.IsNullOrWhiteSpace(s)).ToArray();
                return arr.Length > 0 ? string.Join(", ", arr) : "-";
            }

            // Приводит все числа в строке к формату с 1 знаком после запятой с учётом текущей культуры
            static string FormatNumbersToOne(string? s)
            {
                if (string.IsNullOrWhiteSpace(s)) return s ?? "-";
                return Regex.Replace(s, @"([-+]?\d+[.,]\d+)", m =>
                {
                    var inv = m.Groups[1].Value.Replace(',', '.');
                    if (double.TryParse(inv, NumberStyles.Float, CultureInfo.InvariantCulture, out var num))
                        return num.ToString("F1", CultureInfo.CurrentCulture);
                    return m.Value;
                });
            }

            var r = GenVM.Report!;

            var map = new Dictionary<string, string>
            {
                ["/цикл"] = DashIfEmpty(RawVM.SelectedCycleHeader),

                ["/предСПмакс"] = DashIfEmpty(RawVM.Header.MaxNomen?.ToString()),
                ["/предРАСЧмакс"] = DashIfEmpty(RawVM.Header.MaxCalculated?.ToString()),
                ["/предСПотн"] = DashIfEmpty(RawVM.Header.RelNomen?.ToString()),
                ["/предРАСЧотн"] = DashIfEmpty(RawVM.Header.RelCalculated?.ToString()),

                ["/общмакс"] = $"{r.MaxTotal.Value:F1}",
                ["/общмаксId"] = JoinOrDash(r.MaxTotal.Ids),

                ["/общэкстр"] = DashIfEmpty(FormatNumbersToOne(r.TotalExtrema)),
                ["/сеттэкстр"] = DashIfEmpty(FormatNumbersToOne(r.SettlExtrema)),

                ["/общэкстрId"] = DashIfEmpty(r.TotalExtremaIds),
                ["/сеттэкстрId"] = DashIfEmpty(r.SettlExtremaIds),

                ["/общср"] = $"{r.AvgTotal:F1}",
                ["/сеттмакс"] = $"{r.MaxSettl.Value:F1}",
                ["/сеттмаксId"] = JoinOrDash(r.MaxSettl.Ids),
                ["/сеттср"] = $"{r.AvgSettl:F1}",

                ["/нетдоступа"] = JoinOrDash(r.NoAccessIds),
                ["/уничтожены"] = JoinOrDash(r.DestroyedIds),
                ["/новые"] = JoinOrDash(r.NewIds),

                ["/общ>сп"] = DashIfEmpty(GenVM.ExceedTotalSpDisplay),
                ["/общ>расч"] = DashIfEmpty(GenVM.ExceedTotalCalcDisplay),
                ["/отн>сп"] = DashIfEmpty(GenVM.ExceedRelSpDisplay),
                ["/отн>расч"] = DashIfEmpty(GenVM.ExceedRelCalcDisplay),
            };

            map["/общмин"] = map["/общэкстр"];
            map["/сеттмин"] = map["/сеттэкстр"];
            map["/общминId"] = map["/общэкстрId"];
            map["/сеттминId"] = map["/сеттэкстрId"];

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
                ws.Cell(r, 5).Value = $"{row.Ratio:F5}";
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
            const string sheetName = "Графики динамики";
            const string tableName = "DynTable";

            var ws = wb.Worksheets.FirstOrDefault(s =>
                         s.Name.Equals(sheetName, System.StringComparison.OrdinalIgnoreCase))
                     ?? wb.AddWorksheet(sheetName);

            var cycles = RawVM?.CurrentCycles?.Keys?.OrderBy(c => c).ToList()
                         ?? new System.Collections.Generic.List<int>();

            var dynVm = new DynamicsGrafficViewModel(RawVM, _dynSvc);

            ws.Cell(1, 1).Value = "Id";
            for (int i = 0; i < cycles.Count; i++)
            {
                int cyc = cycles[i];

                string headerText;
                if (RawVM.CycleHeaders.TryGetValue(cyc, out var rawLabel))
                    headerText = CycleLabelParsing.ExtractDateTail(rawLabel) ?? rawLabel;
                else
                    headerText = $"Cycle {cyc}";

                ws.Cell(1, i + 2).Value = headerText;
            }

            var colByCycle = cycles
                .Select((cycle, idx) => new { cycle, col = idx + 2 })
                .ToDictionary(x => x.cycle, x => x.col);

            int r = 2;
            foreach (var ser in dynVm.Lines)
            {
                ws.Cell(r, 1).Value = ser.Id;

                foreach (var pt in ser.Points)
                {
                    if (!colByCycle.TryGetValue(pt.Cycle, out int col)) continue;

                    var cell = ws.Cell(r, col);
                    cell.Value = pt.Mark;
                }
                r++;
            }

            int lastRow = Math.Max(2, r - 1);
            int lastCol = Math.Max(2, cycles.Count + 1);
            var dataRange = ws.Range(1, 1, lastRow, lastCol);

            ws.Row(1).Style.Font.Bold = true;
            dataRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            ws.Columns(1, lastCol).AdjustToContents();

            var dynTable = ws.Tables.FirstOrDefault(t =>
                t.Name.Equals(tableName, System.StringComparison.OrdinalIgnoreCase));

            if (dynTable != null)
            {
                dynTable.Resize(dataRange);
                dynTable.ShowAutoFilter = false;
            }
            else
            {
                var created = dataRange.CreateTable(tableName);
                created.ShowAutoFilter = false;
                created.Theme = XLTableTheme.TableStyleMedium2;
            }

            var wbDynData = wb.DefinedNames.FirstOrDefault(n =>
                n.Name.Equals("DynData", System.StringComparison.OrdinalIgnoreCase));
            wbDynData?.Delete();

            var wsDynData = ws.DefinedNames.FirstOrDefault(n =>
                n.Name.Equals("DynData", System.StringComparison.OrdinalIgnoreCase));
            wsDynData?.Delete();

            wb.CalculateMode = XLCalculateMode.Auto;
        }
    }
}
