using ClosedXML.Excel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Microsoft.Win32;
using Osadka.Models;
using Osadka.Services;
using Osadka.Views;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
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

        public MainViewModel()
        {
            RawVM = new RawDataViewModel();

            var genSvc = new GeneralReportService();
            var relSvc = new RelativeReportService();

            GenVM = new GeneralReportViewModel(RawVM, genSvc, relSvc); // Report пересчитывается при изменениях данных. :contentReference[oaicite:1]{index=1}
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
                           ?? throw new InvalidOperationException("Невалидный формат проекта");

                vm.Header.CycleNumber = data.Cycle;
                vm.Header.MaxNomen = data.MaxNomen;
                vm.Header.MaxCalculated = data.MaxCalculated;
                vm.Header.RelNomen = data.RelNomen;
                vm.Header.RelCalculated = data.RelCalculated;

                vm.DataRows.Clear();
                foreach (var r in data.DataRows) vm.DataRows.Add(r);

                vm.CoordRows.Clear();
                foreach (var c in data.CoordRows) vm.CoordRows.Add(c);

                // --- без обращения к vm.Objects (двусмысленность). Достаём приватное поле _objects.
                var fld = typeof(RawDataViewModel).GetField("_objects", BindingFlags.Instance | BindingFlags.NonPublic);
                if (fld != null)
                {
                    var dict = fld.GetValue(vm) as Dictionary<int, Dictionary<int, List<MeasurementRow>>>
                               ?? new Dictionary<int, Dictionary<int, List<MeasurementRow>>>();
                    dict.Clear();
                    foreach (var objKv in data.Objects)
                    {
                        dict[objKv.Key] = objKv.Value.ToDictionary(
                            cycleKv => cycleKv.Key,
                            cycleKv => cycleKv.Value.ToList());
                    }
                    fld.SetValue(vm, dict);
                }

                vm.ObjectNumbers.Clear();
                // возьмём ключи из _objects
                var currentObjects = (Dictionary<int, Dictionary<int, List<MeasurementRow>>>)(fld?.GetValue(vm)
                                        ?? new Dictionary<int, Dictionary<int, List<MeasurementRow>>>());
                foreach (var obj in currentObjects.Keys.OrderBy(k => k))
                    vm.ObjectNumbers.Add(obj);

                vm.CycleNumbers.Clear();
                if (currentObjects.TryGetValue(vm.Header.ObjectNumber, out var cycles))
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

            // Берём _objects напрямую (чтобы не ссылаться на двусмысленное свойство Objects)
            var fld = typeof(RawDataViewModel).GetField("_objects", BindingFlags.Instance | BindingFlags.NonPublic);
            var currentObjects = (Dictionary<int, Dictionary<int, List<MeasurementRow>>>)(fld?.GetValue(vm)
                                    ?? new Dictionary<int, Dictionary<int, List<MeasurementRow>>>());

            var data = new ProjectData
            {
                Cycle = vm.Header.CycleNumber,
                MaxNomen = vm.Header.MaxNomen,
                MaxCalculated = vm.Header.MaxCalculated,
                RelNomen = vm.Header.RelNomen,
                RelCalculated = vm.Header.RelCalculated,

                DataRows = vm.DataRows.ToList(),
                CoordRows = vm.CoordRows.ToList(),

                Objects = currentObjects.ToDictionary(
                    objKv => objKv.Key,
                    objKv => objKv.Value.ToDictionary(
                        cycleKv => cycleKv.Key,
                        cycleKv => cycleKv.Value.ToList()))
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
                FileName = $"Отчёт_{DateTime.Now:yyyyMMdd_HHmm}.xlsx"
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
                        foreach (var cell in generalWs.CellsUsed(c => c.DataType == XLDataType.Text))
                        {
                            string t = cell.GetString().Trim();
                            if (t.StartsWith("/") && map.TryGetValue(t, out var val))
                                cell.Value = val;
                        }
                    }
                    else
                    {
                        generalWs.Delete();
                    }

                    if (IncludeRelative)
                        AddRelativeSheet(wb);

                    if (IncludeGraphs)
                        AddDynamicsSheet(wb);
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

                workbooks = excelType.InvokeMember("Workbooks",
                    System.Reflection.BindingFlags.GetProperty, null, app, null);
                wb = workbooks.GetType().InvokeMember("Open",
                    System.Reflection.BindingFlags.InvokeMethod, null, workbooks, new object[] { filePath });

                worksheets = wb.GetType().InvokeMember("Worksheets",
                    System.Reflection.BindingFlags.GetProperty, null, wb, null);
                wsData = worksheets.GetType().InvokeMember("Item",
                    System.Reflection.BindingFlags.GetProperty, null, worksheets, new object[] { dataSheetName });

                listObjects = wsData.GetType().InvokeMember("ListObjects",
                    System.Reflection.BindingFlags.GetProperty, null, wsData, null);
                lo = listObjects.GetType().InvokeMember("Item",
                    System.Reflection.BindingFlags.GetProperty, null, listObjects, new object[] { tableName });

                dataBodyRange = lo.GetType().InvokeMember("DataBodyRange",
                    System.Reflection.BindingFlags.GetProperty, null, lo, null);
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
                    catch { /* ignore */ }
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

        // ---------- Плейсхолдеры под новый общий отчёт (экстремумы) с безопасной рефлексией ----------
        private Dictionary<string, string> BuildPlaceholderMap()
        {
            static string DashIfEmpty(string? s) => string.IsNullOrWhiteSpace(s) ? "-" : s;
            static string JoinOrDash(IEnumerable<string>? ids)
                => (ids is null || !ids.Any()) ? "-" : string.Join(", ", ids);

            var r = (object)GenVM.Report!;

            double? ReadDouble(object src, string propName)
            {
                var p = src.GetType().GetProperty(propName);
                if (p == null) return null;
                var val = p.GetValue(src);
                if (val == null) return null;

                // 1) просто double
                if (val is double d) return d;

                // 2) Extremum { Value, Ids }
                var vp = val.GetType().GetProperty("Value");
                if (vp != null && vp.GetValue(val) is double dv) return dv;

                return null;
            }

            IEnumerable<string>? ReadIds(object src, string baseName)
            {
                var pIds = src.GetType().GetProperty(baseName + "Ids");
                if (pIds?.GetValue(src) is IEnumerable<string> idsA) return idsA;

                var p = src.GetType().GetProperty(baseName);
                var val = p?.GetValue(src);
                var idsProp = val?.GetType().GetProperty("Ids");
                if (idsProp?.GetValue(val) is IEnumerable<string> idsB) return idsB;

                return null;
            }

            double? D(params string[] names) { foreach (var n in names) { var v = ReadDouble(r, n); if (v.HasValue) return v; } return null; }
            IEnumerable<string>? IDS(params string[] names) { foreach (var n in names) { var v = ReadIds(r, n); if (v != null) return v; } return null; }

            var map = new Dictionary<string, string>
            {
                ["/цикл"] = DashIfEmpty(RawVM.SelectedCycleHeader),

                ["/предСПмакс"] = DashIfEmpty(RawVM.Header.MaxNomen?.ToString()),
                ["/предРАСЧмакс"] = DashIfEmpty(RawVM.Header.MaxCalculated?.ToString()),
                ["/предСПотн"] = DashIfEmpty(RawVM.Header.RelNomen?.ToString()),
                ["/предРАСЧотн"] = DashIfEmpty(RawVM.Header.RelCalculated?.ToString()),

                ["/вектормакс"] = D("MaxVector", "MaxTotal")?.ToString("F2") ?? "-",
                ["/вектормаксId"] = JoinOrDash(IDS("MaxVector", "MaxTotal")),
                ["/вектормин"] = D("MinVector", "MinTotal")?.ToString("F2") ?? "-",
                ["/векторминId"] = JoinOrDash(IDS("MinVector", "MinTotal")),

                ["/дельтаXмакс"] = D("MaxDx", "MaxDX", "MaxDeltaX")?.ToString("F2") ?? "-",
                ["/дельтаXмаксId"] = JoinOrDash(IDS("MaxDx", "MaxDX", "MaxDeltaX")),
                ["/дельтаXмин"] = D("MinDx", "MinDX", "MinDeltaX")?.ToString("F2") ?? "-",
                ["/дельтаXминId"] = JoinOrDash(IDS("MinDx", "MinDX", "MinDeltaX")),

                ["/дельтаYмакс"] = D("MaxDy", "MaxDY", "MaxDeltaY")?.ToString("F2") ?? "-",
                ["/дельтаYмаксId"] = JoinOrDash(IDS("MaxDy", "MaxDY", "MaxDeltaY")),
                ["/дельтаYмин"] = D("MinDy", "MinDY", "MinDeltaY")?.ToString("F2") ?? "-",
                ["/дельтаYминId"] = JoinOrDash(IDS("MinDy", "MinDY", "MinDeltaY")),

                ["/дельтаHмакс"] = D("MaxDh", "MaxDeltaH")?.ToString("F2") ?? "-",
                ["/дельтаHмаксId"] = JoinOrDash(IDS("MaxDh", "MaxDeltaH")),
                ["/дельтаHмин"] = D("MinDh", "MinDeltaH")?.ToString("F2") ?? "-",
                ["/дельтаHминId"] = JoinOrDash(IDS("MinDh", "MinDeltaH")),

                ["/нетдоступа"] = GenVM.Report?.NoAccessIds is { } na && na.Any() ? string.Join(", ", na) : "-",
                ["/уничтожены"] = GenVM.Report?.DestroyedIds is { } de && de.Any() ? string.Join(", ", de) : "-",
                ["/новые"] = GenVM.Report?.NewIds is { } nw && nw.Any() ? string.Join(", ", nw) : "-",

                ["/общ>сп"] = string.IsNullOrWhiteSpace(GenVM.ExceedTotalSpDisplay) ? "-" : GenVM.ExceedTotalSpDisplay,
                ["/общ>расч"] = string.IsNullOrWhiteSpace(GenVM.ExceedTotalCalcDisplay) ? "-" : GenVM.ExceedTotalCalcDisplay,
                ["/отн>сп"] = string.IsNullOrWhiteSpace(GenVM.ExceedRelSpDisplay) ? "-" : GenVM.ExceedRelSpDisplay,
                ["/отн>расч"] = string.IsNullOrWhiteSpace(GenVM.ExceedRelCalcDisplay) ? "-" : GenVM.ExceedRelCalcDisplay,
            };
            return map;
        }


        private void AddRelativeSheet(XLWorkbook wb)
        {
            var ws = wb.AddWorksheet("Относительная разность");
            ws.Cell(1, 1).Value = "Точка №1";
            ws.Cell(1, 2).Value = "Точка №2";
            ws.Cell(1, 3).Value = "Расстояние";
            ws.Cell(1, 4).Value = "Высотная разность";
            ws.Cell(1, 5).Value = "Относительная";
            ws.Cell(1, 6).Value = "Максимально допустимый предел";

            // НЕ завязано на RelVM.Rows (у тебя его нет).
            var relSvc = new RelativeReportService();
            double spLim = RawVM.Header.RelNomen ?? 0;
            double calcLim = RawVM.Header.RelCalculated ?? 0;
            var relResult = relSvc.Build(RawVM.CoordRows, RawVM.DataRows, spLim, calcLim);

            // Достаём relResult.Rows через рефлексию
            var rowsProp = relResult.GetType().GetProperty("Rows");
            var rowsObj = rowsProp?.GetValue(relResult) as IEnumerable;
            int r = 2;

            if (rowsObj != null)
            {
                foreach (var row in rowsObj)
                {
                    // ожидаемые имена свойств строки
                    string id1 = GetProp(row, "Id1") ?? "";
                    string id2 = GetProp(row, "Id2") ?? "";
                    double? dist = GetPropD(row, "Distance");
                    double? dh = GetPropD(row, "Dh");
                    double? rel = GetPropD(row, "Relative");
                    double? max = GetPropD(row, "Max");

                    ws.Cell(r, 1).Value = id1;
                    ws.Cell(r, 2).Value = id2;
                    ws.Cell(r, 3).Value = dist ?? 0;
                    ws.Cell(r, 4).Value = dh ?? 0;
                    ws.Cell(r, 5).Value = rel ?? 0;
                    ws.Cell(r, 6).Value = max ?? 0;
                    r++;
                }
            }

            ws.Columns().AdjustToContents();

            static string? GetProp(object obj, string name)
                => obj.GetType().GetProperty(name)?.GetValue(obj)?.ToString();

            static double? GetPropD(object obj, string name)
            {
                var p = obj.GetType().GetProperty(name);
                if (p == null) return null;
                var v = p.GetValue(obj);
                if (v is double d) return d;
                if (v is float f) return (double)f;
                if (v is decimal m) return (double)m;
                if (v is int i) return i;
                if (v is long l) return l;
                if (v is string s && double.TryParse(s, out var ds)) return ds;
                return null;
            }
        }

        private void AddDynamicsSheet(XLWorkbook wb)
        {
            const string sheetName = "Графики динамики";
            const string tableName = "DynTable";

            var ws = wb.Worksheets.FirstOrDefault(s =>
                         s.Name.Equals(sheetName, System.StringComparison.OrdinalIgnoreCase))
                     ?? wb.AddWorksheet(sheetName);

            var cycles = RawVM?.CurrentCycles?.Keys?.OrderBy(c => c).ToList()
                         ?? new List<int>();

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

            int rr = 2;
            foreach (var ser in dynVm.Lines)
            {
                ws.Cell(rr, 1).Value = ser.Id;
                foreach (var pt in ser.Points)
                {
                    if (!colByCycle.TryGetValue(pt.Cycle, out int col)) continue;
                    ws.Cell(rr, col).Value = pt.Mark;
                }
                rr++;
            }

            var lastCol = 1 + cycles.Count;
            var rng = ws.Range(1, 1, rr - 1, lastCol);
            var existingTable = ws.Tables.FirstOrDefault(t =>
                string.Equals(t.Name, tableName, StringComparison.OrdinalIgnoreCase));

            if (existingTable != null)
            {
                // просто растягиваем существующую таблицу на актуальный диапазон
                existingTable.Resize(rng);
            }
            else
            {
                rng.CreateTable(tableName);
            }


            ws.Columns().AdjustToContents();

            // подчистим именованные диапазоны от прошлых запусков
            var wbDynTable = wb.DefinedNames.FirstOrDefault(n =>
                n.Name.Equals(tableName, System.StringComparison.OrdinalIgnoreCase));
            wbDynTable?.Delete();

            var wsDynTable = ws.DefinedNames.FirstOrDefault(n =>
                n.Name.Equals(tableName, System.StringComparison.OrdinalIgnoreCase));
            wsDynTable?.Delete();

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
