// File: ViewModels/MainViewModel.cs
using System.Collections.Generic;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;
using System.Windows;
using ClosedXML.Excel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Microsoft.Win32;
using Osadka.Models;
using Osadka.Services;
using Osadka.Services.Abstractions;
using Osadka.Views;

namespace Osadka.ViewModels
{
    public partial class MainViewModel : ObservableObject
    {
        private string? _currentPath;

        // Services (injected via DI)
        private readonly IMessageBoxService _messageBox;
        private readonly IFileDialogService _fileDialog;
        private readonly IFileService _fileService;
        private readonly IProjectService _projectService;

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
        public IRelayCommand PasteProxyCommand { get; }
        private CoordinateExporting? _coord;

        private readonly HashSet<MeasurementRow> _trackedMeasurementRows = new();
        private readonly HashSet<CoordRow> _trackedCoordRows = new();
        private bool _isDirty;

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

            if (_fileService.FileExists(docx))
                _fileService.OpenInDefaultApp(docx);
            else
                _messageBox.Show("Файл справки не найден.", "Справка");
        }

        public MainViewModel(
            RawDataViewModel rawDataViewModel,
            IMessageBoxService messageBox,
            IFileDialogService fileDialog,
            IFileService fileService,
            IProjectService projectService,
            GeneralReportService generalReportService,
            RelativeReportService relativeReportService,
            DynamicsReportService dynamicsReportService)
        {
            // Inject services
            _messageBox = messageBox;
            _fileDialog = fileDialog;
            _fileService = fileService;
            _projectService = projectService;

            // Inject ViewModels and Services
            RawVM = rawDataViewModel;
            _dynSvc = dynamicsReportService;

            // Setup event subscriptions
            RawVM.PropertyChanged += (_, e) =>
            {
                if (e.PropertyName == nameof(RawDataViewModel.TemplatePath) ||
                    e.PropertyName == nameof(RawDataViewModel.CoordUnit) ||
                    e.PropertyName == nameof(RawDataViewModel.DrawingPath))
                {
                    MarkDirty();
                }
            };

            RawVM.Header.PropertyChanged += Header_PropertyChanged;
            RawVM.DataRows.CollectionChanged += DataRows_CollectionChanged;
            RawVM.CoordRows.CollectionChanged += CoordRows_CollectionChanged;

            foreach (var row in RawVM.DataRows)
                SubscribeMeasurementRow(row);
            foreach (var row in RawVM.CoordRows)
                SubscribeCoordRow(row);

            // Create other ViewModels
            GenVM = new GeneralReportViewModel(RawVM, generalReportService, relativeReportService);
            RelVM = new RelativeSettlementsViewModel(RawVM, relativeReportService);

            HelpCommand = new RelayCommand(OpenHelp);
            NavigateCommand = new RelayCommand<string>(Navigate);
            NewProjectCommand = new RelayCommand(NewProject);
            OpenProjectCommand = new RelayCommand(OpenProject);
            SaveProjectCommand = new RelayCommand(SaveProject, () => _isDirty);
            SaveAsProjectCommand = new RelayCommand(SaveAsProject);
            QuickReportCommand = new RelayCommand(DoQuickExport, () => GenVM.Report != null);
            PasteProxyCommand = new RelayCommand(
                () => RawVM.PasteCommand.Execute(null),
                () => RawVM.PasteCommand.CanExecute(null));

            RawVM.PasteCommand.CanExecuteChanged += (_, _) =>
            {
                if (PasteProxyCommand is RelayCommand pasteRelay)
                    pasteRelay.NotifyCanExecuteChanged();
            };

            GenVM.PropertyChanged += (_, e) =>
            {
                if (e.PropertyName == nameof(GeneralReportViewModel.Report) &&
                    QuickReportCommand is RelayCommand quickRelay)
                {
                    quickRelay.NotifyCanExecuteChanged();
                }
            };

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
            ResetDirty();
        }

        public void LoadProject(string path)
        {
            if (RawVM is not { } vm) return;

            try
            {
                var (data, dwgPath) = _projectService.Load(path);

                vm.Header.CycleNumber = data.Cycle;
                vm.Header.MaxNomen = data.MaxNomen;
                vm.Header.MaxCalculated = data.MaxCalculated;
                vm.Header.RelNomen = data.RelNomen;
                vm.Header.RelCalculated = data.RelCalculated;
                vm.SelectedCycleHeader = data.SelectedCycleHeader ?? string.Empty;

                vm.DataRows.Clear();
                foreach (var r in data.DataRows) vm.DataRows.Add(r);

                vm.CoordRows.Clear();
                foreach (var r in data.CoordRows) vm.CoordRows.Add(r);

                vm.Objects.Clear();
                foreach (var obj in data.Objects)
                    vm.Objects[obj.Key] = obj.Value.ToDictionary(kv => kv.Key, kv => kv.Value.ToList());

                vm.DrawingPath = dwgPath;

                _currentPath = path;
                ResetDirty();
            }
            catch (Exception ex)
            {
                _messageBox.ShowWithOptions(
                    $"Ошибка при загрузке проекта:\n{ex.Message}",
                    "Ошибка",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
        }



        private void OpenProject()
        {
            var path = _fileDialog.OpenFile("Osadka Project (*.osd)|*.osd|All Files|*.*");
            if (path == null) return;

            LoadProject(path);
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
            var path = _fileDialog.SaveFile("Osadka Project (*.osd)|*.osd");
            if (path == null) return;

            SaveTo(path);
            _currentPath = path;
        }

        void SaveTo(string path)
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

            _projectService.Save(path, data, vm.DrawingPath);
            ResetDirty();
        }



        private void SubscribeMeasurementRow(MeasurementRow? row)
        {
            if (row is null) return;
            if (_trackedMeasurementRows.Add(row))
            {
                row.PropertyChanged += MeasurementRow_PropertyChanged;
            }
        }

        private void UnsubscribeMeasurementRow(MeasurementRow? row)
        {
            if (row is null) return;
            if (_trackedMeasurementRows.Remove(row))
            {
                row.PropertyChanged -= MeasurementRow_PropertyChanged;
            }
        }

        private void SubscribeCoordRow(CoordRow? row)
        {
            if (row is null) return;
            if (_trackedCoordRows.Add(row))
            {
                row.PropertyChanged += CoordRow_PropertyChanged;
            }
        }

        private void UnsubscribeCoordRow(CoordRow? row)
        {
            if (row is null) return;
            if (_trackedCoordRows.Remove(row))
            {
                row.PropertyChanged -= CoordRow_PropertyChanged;
            }
        }

        private void DataRows_CollectionChanged(object? sender, NotifyCollectionChangedEventArgs e)
        {
            if (e.Action == NotifyCollectionChangedAction.Reset)
            {
                ClearMeasurementRowSubscriptions();
                foreach (var row in RawVM.DataRows)
                    SubscribeMeasurementRow(row);
            }
            else
            {
                if (e.OldItems != null)
                    foreach (MeasurementRow row in e.OldItems)
                        UnsubscribeMeasurementRow(row);

                if (e.NewItems != null)
                    foreach (MeasurementRow row in e.NewItems)
                        SubscribeMeasurementRow(row);
            }

            MarkDirty();
        }

        private void CoordRows_CollectionChanged(object? sender, NotifyCollectionChangedEventArgs e)
        {
            if (e.Action == NotifyCollectionChangedAction.Reset)
            {
                ClearCoordRowSubscriptions();
                foreach (var row in RawVM.CoordRows)
                    SubscribeCoordRow(row);
            }
            else
            {
                if (e.OldItems != null)
                    foreach (CoordRow row in e.OldItems)
                        UnsubscribeCoordRow(row);

                if (e.NewItems != null)
                    foreach (CoordRow row in e.NewItems)
                        SubscribeCoordRow(row);
            }

            MarkDirty();
        }

        private void MeasurementRow_PropertyChanged(object? sender, PropertyChangedEventArgs e)
            => MarkDirty();

        private void CoordRow_PropertyChanged(object? sender, PropertyChangedEventArgs e)
            => MarkDirty();

        private void Header_PropertyChanged(object? sender, PropertyChangedEventArgs e)
            => MarkDirty();

        private void ClearMeasurementRowSubscriptions()
        {
            foreach (var row in _trackedMeasurementRows.ToList())
                row.PropertyChanged -= MeasurementRow_PropertyChanged;
            _trackedMeasurementRows.Clear();
        }

        private void ClearCoordRowSubscriptions()
        {
            foreach (var row in _trackedCoordRows.ToList())
                row.PropertyChanged -= CoordRow_PropertyChanged;
            _trackedCoordRows.Clear();
        }

        private void MarkDirty()
        {
            if (_isDirty)
                return;

            _isDirty = true;
            if (SaveProjectCommand is RelayCommand saveRelay)
                saveRelay.NotifyCanExecuteChanged();
        }

        private void ResetDirty()
        {
            if (!_isDirty)
            {
                if (SaveProjectCommand is RelayCommand saveRelay)
                    saveRelay.NotifyCanExecuteChanged();
                return;
            }

            _isDirty = false;
            if (SaveProjectCommand is RelayCommand relay)
                relay.NotifyCanExecuteChanged();
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

            // Алиасы старых тегов
            map["/общмин"] = map["/общэкстр"];
            map["/сеттмин"] = map["/сеттэкстр"];
            map["/общминId"] = map["/общэкстрId"];
            map["/сеттминId"] = map["/сеттэкстрId"];

            // === НОВОЕ: читаем максимум относительной разницы из бизнес-логики Relative ===
            var mr = RelVM?.Report?.MaxRelative;
            if (mr is { } && !double.IsNaN(mr.Value))
            {
                map["/отнмакс"] = mr.Value.ToString("F5", CultureInfo.CurrentCulture);
                map["/отнмаксId"] = JoinOrDash(mr.Ids);
            }
            else
            {
                map["/отнмакс"] = "-";
                map["/отнмаксId"] = "-";
            }

            return map;
        }

        private void AddRelativeSheet(XLWorkbook wb)
        {
            // Получаем/создаём лист
            var ws = wb.Worksheets.FirstOrDefault(s =>
                         s.Name.Equals("Относительная разность", StringComparison.OrdinalIgnoreCase))
                     ?? wb.AddWorksheet("Относительная разность");
            ws.Clear();

            // Заголовки (как на скрине)
            ws.Cell(1, 1).Value = "№1";
            ws.Cell(1, 2).Value = "№2";
            ws.Cell(1, 3).Value = "Dist, мм";
            ws.Cell(1, 4).Value = "ΔS, мм";
            ws.Cell(1, 5).Value = "ΔS/Dist";

            // Хелпер: число или прочерк
            static void SetNumberOrDash(IXLCell cell, double value, string? format = null)
            {
                if (double.IsNaN(value) || double.IsInfinity(value))
                {
                    cell.Value = "-";
                   // cell.XLDataType(XLDataType.Text);           // вместо cell.DataType = ...
                    cell.Style.NumberFormat.Format = "@";        // принудительно "Текст"
                }
                else
                {
                    cell.Value = value;
                    if (!string.IsNullOrWhiteSpace(format))
                        cell.Style.NumberFormat.Format = format;
                }
            }

            // Данные
            int r = 2;
            foreach (var row in RelVM.AllRows)
            {
                ws.Cell(r, 1).Value = row.Id1;
                ws.Cell(r, 2).Value = row.Id2;

                // Форматы при необходимости можно подправить
                SetNumberOrDash(ws.Cell(r, 3), row.Distance, "0.0");      // мм
                SetNumberOrDash(ws.Cell(r, 4), row.DeltaTotal, "0.0");      // мм
                SetNumberOrDash(ws.Cell(r, 5), row.Ratio, "0.000000"); // безразмерный

                r++;
            }

            // Оформление
            ws.Range(1, 1, 1, 5).Style.Font.Bold = true;
            ws.Columns(1, 5).AdjustToContents();
            ws.SheetView.FreezeRows(1);
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
            var used = new Dictionary<string, int>(StringComparer.CurrentCultureIgnoreCase);
            string Unique(string h)
            {
                if (string.IsNullOrWhiteSpace(h)) h = "Cycle";
                if (used.TryGetValue(h, out var n))
                {
                    n++;
                    used[h] = n;
                    return $"{h} #{n}";
                }
                used[h] = 1;
                return h;
            }
            ws.Cell(1, 1).Value = "Id";
            for (int i = 0; i < cycles.Count; i++)
            {
                int cyc = cycles[i];
                string headerText;
                if (RawVM.CycleHeaders.TryGetValue(cyc, out var rawLabel))
                    headerText = CycleLabelParsing.ExtractDateTail(rawLabel) ?? rawLabel;
                else
                    headerText = $"Cycle {cyc}";

                // NEW: исключаем дубликаты
                ws.Cell(1, i + 2).Value = Unique(headerText.Trim());
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
