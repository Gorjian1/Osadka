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
using Osadka.Models.Cycles;
using Osadka.Services.Reports;
using Osadka.Services.Data;
using Osadka.Services.Parsing;
using Osadka.Views;
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
        public CycleGroupsViewModel CycleGroupsVM => _cycleGroupsViewModel ??= new CycleGroupsViewModel(RawVM);
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
        private CycleGroupsPage? _cycleGroupsPage;
        private CycleGroupsViewModel? _cycleGroupsViewModel;

        private readonly HashSet<MeasurementRow> _trackedMeasurementRows = new();
        private readonly HashSet<CoordRow> _trackedCoordRows = new();
        private bool _isDirty;
        private bool _suppressDirtyTracking; // Флаг для подавления отслеживания во время загрузки

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
            public const string Scale = "CycleScale";
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
            // Подавляем отслеживание изменений во время инициализации
            _suppressDirtyTracking = true;

            RawVM = new RawDataViewModel();
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
            RawVM.ActiveFilterChanged += (_, __) => MarkDirty();

            foreach (var row in RawVM.DataRows)
                SubscribeMeasurementRow(row);
            foreach (var row in RawVM.CoordRows)
                SubscribeCoordRow(row);

            var genSvc = new GeneralReportService();
            var relSvc = new RelativeReportService();

            GenVM = new GeneralReportViewModel(RawVM, genSvc, relSvc);
            RelVM = new RelativeSettlementsViewModel(RawVM, relSvc);
            _dynSvc = new DynamicsReportService();

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

            // Снимаем подавление - теперь отслеживаем изменения
            _suppressDirtyTracking = false;
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
                PageKeys.Scale => _cycleGroupsPage ??= new CycleGroupsPage(CycleGroupsVM),
                PageKeys.Graf => new DynamicsGrafficPage(new DynamicsGrafficViewModel(RawVM, _dynSvc)),
                _ => CurrentPage
            };
        }

        #endregion

        private void NewProject()
        {
            _suppressDirtyTracking = true;
            try
            {
                RawVM.ClearCommand.Execute(null);
                _currentPath = null;
                ResetDirty();
            }
            finally
            {
                _suppressDirtyTracking = false;
            }
        }

        public void LoadProject(string path)
        {
            if (RawVM is not { } vm) return;

            _suppressDirtyTracking = true;
            try
            {
                var json = System.IO.File.ReadAllText(path);

                using var doc = System.Text.Json.JsonDocument.Parse(json);
                string? dwgPath = doc.RootElement.TryGetProperty("DwgPath", out var p) ? p.GetString() : null;

                var data = System.Text.Json.JsonSerializer.Deserialize<ProjectData>(json)
                           ?? throw new InvalidOperationException("Невалидный формат");

                vm.SuspendRefresh(true);
                try
                {
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
                }
                finally
                {
                    vm.SuspendRefresh(false);
                }

                vm.ObjectNumbers.Clear();
                foreach (var k in vm.Objects.Keys.OrderBy(k => k)) vm.ObjectNumbers.Add(k);

                int objectNumber = data.ObjectNumber;
                if (!vm.ObjectNumbers.Contains(objectNumber))
                    objectNumber = vm.ObjectNumbers.FirstOrDefault();
                if (objectNumber == 0)
                    objectNumber = 1;
                vm.Header.ObjectNumber = objectNumber;

                vm.CycleNumbers.Clear();
                if (vm.Objects.TryGetValue(objectNumber, out var cyclesForObject))
                {
                    foreach (var k in cyclesForObject.Keys.OrderBy(k => k)) vm.CycleNumbers.Add(k);
                }

                int cycleNumber = data.Cycle;
                if (!vm.CycleNumbers.Contains(cycleNumber))
                    cycleNumber = vm.CycleNumbers.FirstOrDefault();
                if (cycleNumber == 0)
                    cycleNumber = 1;
                vm.Header.CycleNumber = cycleNumber;

                vm.SelectedCycleHeader = data.SelectedCycleHeader ?? string.Empty;

                try
                {
                    vm.RebuildCycleGroups();
                    vm.SetDisabledPoints(data.DisabledPointIds?.ToArray() ?? Array.Empty<string>());
                }
                catch (Exception rebuildEx)
                {
                    // Если RebuildCycleGroups не удался, логируем ошибку но продолжаем загрузку
                    System.Diagnostics.Debug.WriteLine($"Warning: RebuildCycleGroups failed: {rebuildEx.Message}");
                    // Очищаем группы циклов, чтобы избежать некорректного состояния
                    vm.CycleGroups.Clear();
                }

                vm.DrawingPath = dwgPath;

                _currentPath = path;
                ResetDirty();
            }
            catch (Exception ex)
            {
                // При критической ошибке очищаем состояние, чтобы не оставить приложение в поврежденном виде
                vm.SuspendRefresh(true);
                try
                {
                    vm.Objects.Clear();
                    vm.ObjectNumbers.Clear();
                    vm.CycleNumbers.Clear();
                    vm.CycleGroups.Clear();
                    vm.DataRows.Clear();
                    vm.CoordRows.Clear();
                }
                finally
                {
                    vm.SuspendRefresh(false);
                }

                System.Windows.MessageBox.Show(
                    $"Ошибка при загрузке проекта:\n{ex.Message}",
                    "Ошибка",
                    System.Windows.MessageBoxButton.OK,
                    System.Windows.MessageBoxImage.Error
                );
            }
            finally
            {
                _suppressDirtyTracking = false;
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
                ObjectNumber = vm.Header.ObjectNumber,
                DataRows = vm.DataRows.ToList(),
                CoordRows = vm.CoordRows.ToList(),
                Objects = vm.Objects.ToDictionary(
                    objKv => objKv.Key,
                    objKv => objKv.Value.ToDictionary(
                        cycleKv => cycleKv.Key,
                        cycleKv => cycleKv.Value.ToList()
                    )),
                DisabledPointIds = new HashSet<string>(vm.DisabledPoints, StringComparer.OrdinalIgnoreCase)
            };
            var node = JsonSerializer.SerializeToNode(data)!.AsObject();
            node["DwgPath"] = vm.DrawingPath;

            File.WriteAllText(
                path,
                node.ToJsonString(new JsonSerializerOptions { WriteIndented = true })
            );

            ResetDirty();
        }

        /// <summary>
        /// Обработка закрытия окна. Проверяет наличие несохраненных изменений.
        /// </summary>
        /// <returns>true если можно закрывать окно, false если отменить закрытие</returns>
        public bool OnWindowClosing()
        {
            if (!_isDirty)
                return true; // Нет изменений - можно закрывать

            var result = MessageBox.Show(
                "Имеются несохраненные изменения. Сохранить проект перед выходом?",
                "Несохраненные изменения",
                MessageBoxButton.YesNoCancel,
                MessageBoxImage.Question);

            switch (result)
            {
                case MessageBoxResult.Yes:
                    // Сохранить и закрыть
                    if (_currentPath != null)
                    {
                        SaveTo(_currentPath);
                        return true;
                    }
                    else
                    {
                        // Нет пути - показываем диалог SaveAs
                        var dlg = new SaveFileDialog
                        {
                            Filter = "Osadka Project (*.osd)|*.osd"
                        };
                        if (dlg.ShowDialog() == true)
                        {
                            SaveTo(dlg.FileName);
                            _currentPath = dlg.FileName;
                            return true;
                        }
                        // Пользователь отменил сохранение - не закрываем
                        return false;
                    }

                case MessageBoxResult.No:
                    // Закрыть без сохранения
                    return true;

                case MessageBoxResult.Cancel:
                default:
                    // Отменить закрытие
                    return false;
            }
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
            // Игнорируем во время загрузки/инициализации
            if (_suppressDirtyTracking)
                return;

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

                // Подготовка данных для экспорта
                var exportData = new ExcelExportData
                {
                    PlaceholderMap = BuildPlaceholderMap(),
                    DisabledTags = GenVM.Settings?.GetDisabledTags() ?? new HashSet<string>(),
                    RelativeRows = RelVM.AllRows.Select(row => new RelativeSettlementRow
                    {
                        Id1 = row.Id1,
                        Id2 = row.Id2,
                        Distance = row.Distance,
                        DeltaTotal = row.DeltaTotal,
                        Ratio = row.Ratio
                    }).ToList(),
                    ActiveCycles = RawVM?.GetActiveCyclesSnapshot() ?? new Dictionary<int, List<MeasurementRow>>(),
                    CycleHeaders = RawVM?.CycleHeaders != null
                        ? new Dictionary<int, string>(RawVM.CycleHeaders)
                        : new Dictionary<int, string>(),
                    DynamicsData = _dynSvc.Build(RawVM?.GetActiveCyclesSnapshot() ?? new Dictionary<int, List<MeasurementRow>>())
                        .Select(s => new DynamicsSeries
                        {
                            Id = s.Id,
                            Points = s.Points.Select(p => new DynamicsPoint
                            {
                                Cycle = p.Cycle,
                                Mark = p.Mark
                            }).ToList()
                        }).ToList()
                };

                var options = new ExcelExportOptions
                {
                    IncludeGeneral = IncludeGeneral,
                    IncludeRelative = IncludeRelative,
                    IncludeGraphs = IncludeGraphs,
                    TemplatePath = template,
                    OutputPath = dlg.FileName
                };

                var exportService = new ExcelExportService();
                exportService.ExportToExcel(options, exportData);

                if (IncludeGraphs)
                {
                    RunSta(() => exportService.BuildExcelChart(
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

    }
}
