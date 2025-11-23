using ClosedXML.Excel;
using Osadka.Services.Abstractions;
using Osadka.ViewModels;
using Serilog;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Windows;

namespace Osadka.Services.Implementation;

/// <summary>
/// Реализация сервиса экспорта в Excel
/// </summary>
public class ExcelExportService : IExportService
{
    private readonly IMessageBoxService _messageBox;
    private readonly IFileDialogService _fileDialog;
    private readonly DynamicsReportService _dynamicsService;

    public ExcelExportService(
        IMessageBoxService messageBox,
        IFileDialogService fileDialog,
        DynamicsReportService dynamicsService)
    {
        _messageBox = messageBox;
        _fileDialog = fileDialog;
        _dynamicsService = dynamicsService;
    }

    public void QuickExport(
        RawDataViewModel rawDataViewModel,
        GeneralReportViewModel generalReportViewModel,
        RelativeSettlementsViewModel relativeSettlementsViewModel,
        bool includeGeneral,
        bool includeRelative,
        bool includeGraphs)
    {
        if (generalReportViewModel.Report is null)
        {
            _messageBox.Show("Отчёт не готов для экспорта", "Экспорт");
            return;
        }

        if (!(includeGeneral || includeRelative || includeGraphs))
        {
            _messageBox.Show(
                "Выберите хотя бы один пункт: Общий, Относительный, Графики.",
                "Экспорт");
            return;
        }

        // Выбор шаблона: сначала пользовательский, затем встроенный
        string? userTemplate = rawDataViewModel?.TemplatePath;
        string exeDir = AppContext.BaseDirectory;
        string fallbackTemplate = Path.Combine(exeDir, "template.xlsx");
        string template = (!string.IsNullOrWhiteSpace(userTemplate) && File.Exists(userTemplate))
            ? userTemplate!
            : fallbackTemplate;

        if (!File.Exists(template))
        {
            _messageBox.ShowWithOptions(
                "Шаблон Excel не найден: ни выбранный, ни встроенный template.xlsx.",
                "Экспорт",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
            return;
        }

        // Подстраиваем расширение под шаблон (.xlsx или .xlsm)
        string ext = Path.GetExtension(template).Equals(".xlsm", StringComparison.OrdinalIgnoreCase)
            ? "*.xlsm"
            : "*.xlsx";
        string filter = $"Excel ({ext})|{ext}";

        var outputPath = _fileDialog.SaveFile(
            filter,
            $"Отчёт_{DateTime.Now:yyyyMMdd_HHmm}{Path.GetExtension(template)}");

        if (outputPath == null) return;

        try
        {
            File.Copy(template, outputPath, overwrite: true);

            using (var wb = new XLWorkbook(outputPath))
            {
                var generalWs = wb.Worksheets.First(); // титульный/общий

                if (includeGeneral)
                {
                    var map = BuildPlaceholderMap(
                        rawDataViewModel,
                        generalReportViewModel,
                        relativeSettlementsViewModel);

                    // ВАЖНО: поддержка выключенных блоков — удаляем строки, где встретились их теги
                    var disabled = generalReportViewModel.Settings?.GetDisabledTags()
                                   ?? new HashSet<string>();

                    // Снимок текстовых ячеек, чтобы удалять строки уже после прохода
                    var textCells = generalWs.CellsUsed(c => c.DataType == XLDataType.Text).ToList();
                    var rowsToDelete = new HashSet<int>();

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

                if (includeRelative)
                    AddRelativeSheet(wb, relativeSettlementsViewModel);

                if (includeGraphs)
                {
                    AddDynamicsSheet(wb, rawDataViewModel, _dynamicsService);
                }
                else
                {
                    var dynWs = wb.Worksheets.FirstOrDefault(
                        ws => string.Equals(ws.Name, "Графики динамики", StringComparison.OrdinalIgnoreCase));
                    dynWs?.Delete();
                }

                wb.Save();
            }

            if (includeGraphs)
            {
                try
                {
                    Log.Information("Начало создания графиков через COM Interop для файла {FilePath}", outputPath);

                    if (!IsExcelAvailable())
                    {
                        Log.Warning("Microsoft Excel недоступен для создания графиков");
                        var result = _messageBox.ShowWithOptions(
                            "Microsoft Excel недоступен.\n\n" +
                            "Графики не могут быть созданы, но отчёт сохранён без графиков.\n\n" +
                            "Для создания графиков необходим установленный Microsoft Office Excel.",
                            "Предупреждение",
                            MessageBoxButton.OK,
                            MessageBoxImage.Warning);
                    }
                    else
                    {
                        RunSta(() => BuildChartFromDynTable_Quick_NoPIA(
                            filePath: outputPath,
                            dataSheetName: "Графики динамики",
                            tableName: "DynTable",
                            chartSheetName: "Графики динамики",
                            left: 40, top: 200, width: 920, height: 440,
                            deleteOldCharts: true
                        ));

                        Log.Information("Графики успешно созданы");
                    }
                }
                catch (Exception chartEx)
                {
                    Log.Error(chartEx, "Ошибка при создании графиков через COM Interop");

                    var result = _messageBox.ShowWithOptions(
                        $"Ошибка создания графиков:\n{chartEx.Message}\n\n" +
                        "Отчёт сохранён без графиков.\n\n" +
                        $"Детали ошибки сохранены в логах.\nТип ошибки: {chartEx.GetType().Name}",
                        "Предупреждение",
                        MessageBoxButton.OK,
                        MessageBoxImage.Warning);
                }
            }

            _messageBox.Show("Экспорт завершён", "Экспорт");
        }
        catch (Exception ex)
        {
            Log.Error(ex, "Критическая ошибка экспорта отчёта");

            _messageBox.ShowWithOptions(
                $"Ошибка экспорта: {ex.Message}\n\n" +
                $"Тип: {ex.GetType().Name}\n" +
                $"Логи: %AppData%\\Osadka\\logs",
                "Экспорт",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
        }
    }

    private Dictionary<string, string> BuildPlaceholderMap(
        RawDataViewModel rawDataViewModel,
        GeneralReportViewModel generalReportViewModel,
        RelativeSettlementsViewModel relativeSettlementsViewModel)
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

        var r = generalReportViewModel.Report!;

        var map = new Dictionary<string, string>
        {
            ["/цикл"] = DashIfEmpty(rawDataViewModel.SelectedCycleHeader),

            ["/предСПмакс"] = DashIfEmpty(rawDataViewModel.Header.MaxNomen?.ToString()),
            ["/предРАСЧмакс"] = DashIfEmpty(rawDataViewModel.Header.MaxCalculated?.ToString()),
            ["/предСПотн"] = DashIfEmpty(rawDataViewModel.Header.RelNomen?.ToString()),
            ["/предРАСЧотн"] = DashIfEmpty(rawDataViewModel.Header.RelCalculated?.ToString()),

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

            ["/общ>сп"] = DashIfEmpty(generalReportViewModel.ExceedTotalSpDisplay),
            ["/общ>расч"] = DashIfEmpty(generalReportViewModel.ExceedTotalCalcDisplay),
            ["/отн>сп"] = DashIfEmpty(generalReportViewModel.ExceedRelSpDisplay),
            ["/отн>расч"] = DashIfEmpty(generalReportViewModel.ExceedRelCalcDisplay),
        };

        // Алиасы старых тегов
        map["/общмин"] = map["/общэкстр"];
        map["/сеттмин"] = map["/сеттэкстр"];
        map["/общминId"] = map["/общэкстрId"];
        map["/сеттминId"] = map["/сеттэкстрId"];

        // === НОВОЕ: читаем максимум относительной разницы из бизнес-логики Relative ===
        var mr = relativeSettlementsViewModel?.Report?.MaxRelative;
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

    private void AddRelativeSheet(XLWorkbook wb, RelativeSettlementsViewModel relativeSettlementsViewModel)
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
        foreach (var row in relativeSettlementsViewModel.AllRows)
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

    private void AddDynamicsSheet(
        XLWorkbook wb,
        RawDataViewModel rawDataViewModel,
        DynamicsReportService dynamicsService)
    {
        const string sheetName = "Графики динамики";
        const string tableName = "DynTable";

        var ws = wb.Worksheets.FirstOrDefault(s =>
                     s.Name.Equals(sheetName, StringComparison.OrdinalIgnoreCase))
                 ?? wb.AddWorksheet(sheetName);

        var cycles = rawDataViewModel?.CurrentCycles?.Keys?.OrderBy(c => c).ToList()
                     ?? new List<int>();

        var dynVm = new DynamicsGrafficViewModel(rawDataViewModel, dynamicsService);
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
            if (rawDataViewModel.CycleHeaders.TryGetValue(cyc, out var rawLabel))
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
            t.Name.Equals(tableName, StringComparison.OrdinalIgnoreCase));

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
            n.Name.Equals("DynData", StringComparison.OrdinalIgnoreCase));
        wbDynData?.Delete();

        var wsDynData = ws.DefinedNames.FirstOrDefault(n =>
            n.Name.Equals("DynData", StringComparison.OrdinalIgnoreCase));
        wsDynData?.Delete();

        wb.CalculateMode = XLCalculateMode.Auto;
    }

    /// <summary>
    /// Проверяет доступность Microsoft Excel через COM
    /// </summary>
    private static bool IsExcelAvailable()
    {
        try
        {
            var excelType = Type.GetTypeFromProgID("Excel.Application", throwOnError: false);
            if (excelType == null)
            {
                Log.Warning("Excel.Application не зарегистрирован в COM");
                return false;
            }

            // Пытаемся создать экземпляр
            object? app = Activator.CreateInstance(excelType);
            if (app == null)
            {
                Log.Warning("Не удалось создать экземпляр Excel.Application");
                return false;
            }

            // Закрываем созданный экземпляр
            excelType.InvokeMember("Quit", System.Reflection.BindingFlags.InvokeMethod, null, app, null);
            if (Marshal.IsComObject(app))
                Marshal.ReleaseComObject(app);

            Log.Debug("Excel доступен для COM Automation");
            return true;
        }
        catch (Exception ex)
        {
            Log.Warning(ex, "Ошибка при проверке доступности Excel");
            return false;
        }
    }

    private static void RunSta(Action action)
    {
        Exception? thrownException = null;

        var t = new System.Threading.Thread(() =>
        {
            try
            {
                action();
            }
            catch (Exception ex)
            {
                thrownException = ex;
                Log.Error(ex, "Ошибка в STA потоке при выполнении COM операции");
            }
        })
        { IsBackground = true };

        t.SetApartmentState(System.Threading.ApartmentState.STA);
        t.Start();
        t.Join();

        // Пробрасываем исключение из потока
        if (thrownException != null)
            throw new InvalidOperationException("Ошибка в COM Interop потоке", thrownException);
    }

    private static void BuildChartFromDynTable_Quick_NoPIA(
        string filePath,
        string dataSheetName = "Графики динамики",
        string tableName = "DynTable",
        string chartSheetName = "Графики динамики",
        int left = 40, int top = 200, int width = 920, int height = 440,
        bool deleteOldCharts = true)
    {
        Log.Debug("BuildChartFromDynTable_Quick_NoPIA: Начало создания графика для файла {FilePath}", filePath);
        Log.Debug("Параметры: dataSheet={DataSheet}, table={Table}, chartSheet={ChartSheet}",
            dataSheetName, tableName, chartSheetName);

        const int xlRows = 1;
        const int xlLine = 4;

        object app = null, wb = null, wsData = null, wsChart = null;
        object workbooks = null, worksheets = null;
        object listObjects = null, lo = null, loRange = null, dataBodyRange = null;
        object chartObjects = null, chartObj = null, chart = null;

        try
        {
            // Шаг 1: Получение типа Excel.Application
            Log.Debug("Шаг 1: Получение типа Excel.Application через ProgID");
            var excelType = Type.GetTypeFromProgID("Excel.Application", throwOnError: false);
            if (excelType == null)
            {
                Log.Error("Excel.Application ProgID не найден - Excel не зарегистрирован в COM");
                throw new InvalidOperationException("Excel.Application не зарегистрирован");
            }
            Log.Debug("Excel.Application тип получен: {TypeName}", excelType.FullName);

            // Шаг 2: Создание экземпляра Excel
            Log.Debug("Шаг 2: Создание экземпляра Excel.Application");
            app = Activator.CreateInstance(excelType);
            if (app == null)
            {
                Log.Error("Не удалось создать экземпляр Excel.Application");
                throw new InvalidOperationException("Не удалось создать экземпляр Excel");
            }
            Log.Debug("Excel.Application экземпляр создан успешно");

            // Шаг 3: Настройка свойств Excel
            Log.Debug("Шаг 3: Настройка Visible=false и DisplayAlerts=false");
            excelType.InvokeMember("Visible", System.Reflection.BindingFlags.SetProperty,
                null, app, new object[] { false });
            excelType.InvokeMember("DisplayAlerts", System.Reflection.BindingFlags.SetProperty,
                null, app, new object[] { false });
            Log.Debug("Свойства Excel настроены");

            // Шаг 4: Открытие workbook
            Log.Debug("Шаг 4: Открытие файла {FilePath}", filePath);
            workbooks = excelType.InvokeMember("Workbooks", System.Reflection.BindingFlags.GetProperty,
                null, app, null);
            var workbooksType = workbooks.GetType();

            wb = workbooksType.InvokeMember("Open", System.Reflection.BindingFlags.InvokeMethod,
                null, workbooks, new object[] { filePath });
            if (wb == null)
            {
                Log.Error("Не удалось открыть файл {FilePath}", filePath);
                throw new InvalidOperationException($"Не удалось открыть файл: {filePath}");
            }
            var wbType = wb.GetType();
            Log.Debug("Workbook открыт успешно");

            // Шаг 5: Получение листов
            Log.Debug("Шаг 5: Получение коллекции Worksheets");
            worksheets = wbType.InvokeMember("Worksheets", System.Reflection.BindingFlags.GetProperty,
                null, wb, null);
            var worksheetsType = worksheets.GetType();
            Log.Debug("Worksheets получены");

            // Шаг 6: Получение листа с данными
            Log.Debug("Шаг 6: Получение листа данных '{DataSheet}'", dataSheetName);
            wsData = worksheetsType.InvokeMember("Item", System.Reflection.BindingFlags.GetProperty,
                null, worksheets, new object[] { dataSheetName });
            if (wsData == null)
            {
                Log.Error("Лист '{DataSheet}' не найден", dataSheetName);
                throw new InvalidOperationException($"Лист '{dataSheetName}' не найден");
            }
            var wsDataType = wsData.GetType();
            Log.Debug("Лист данных получен");

            // Шаг 7: Получение листа для графиков
            Log.Debug("Шаг 7: Получение листа графиков '{ChartSheet}'", chartSheetName);
            wsChart = worksheetsType.InvokeMember("Item", System.Reflection.BindingFlags.GetProperty,
                null, worksheets, new object[] { chartSheetName });
            if (wsChart == null)
            {
                Log.Error("Лист '{ChartSheet}' не найден", chartSheetName);
                throw new InvalidOperationException($"Лист '{chartSheetName}' не найден");
            }
            var wsChartType = wsChart.GetType();
            Log.Debug("Лист графиков получен");

            // Шаг 8: Получение таблицы ListObject
            Log.Debug("Шаг 8: Получение ListObjects из листа данных");
            listObjects = wsDataType.InvokeMember("ListObjects", System.Reflection.BindingFlags.GetProperty,
                null, wsData, null);
            var listObjectsType = listObjects.GetType();

            Log.Debug("Шаг 9: Получение таблицы '{TableName}'", tableName);
            lo = listObjectsType.InvokeMember("Item", System.Reflection.BindingFlags.GetProperty,
                null, listObjects, new object[] { tableName });
            if (lo == null)
            {
                Log.Error("Таблица '{TableName}' не найдена на листе '{DataSheet}'", tableName, dataSheetName);
                throw new InvalidOperationException($"Таблица '{tableName}' не найдена");
            }
            var loType = lo.GetType();
            Log.Debug("Таблица найдена");

            // Шаг 10: Получение диапазонов таблицы
            Log.Debug("Шаг 10: Получение Range и DataBodyRange");
            loRange = loType.InvokeMember("Range", System.Reflection.BindingFlags.GetProperty,
                null, lo, null);
            var loRangeType = loRange.GetType();

            dataBodyRange = loType.InvokeMember("DataBodyRange", System.Reflection.BindingFlags.GetProperty,
                null, lo, null);
            if (dataBodyRange == null)
            {
                Log.Error("DataBodyRange пустой - таблица '{TableName}' не содержит данных", tableName);
                throw new InvalidOperationException($"Таблица '{tableName}' не содержит данных");
            }
            Log.Debug("Диапазоны данных получены");

            // Шаг 11: Удаление старых графиков
            if (deleteOldCharts)
            {
                Log.Debug("Шаг 11: Удаление старых графиков с листа '{ChartSheet}'", chartSheetName);
                chartObjects = wsChartType.InvokeMember("ChartObjects", System.Reflection.BindingFlags.InvokeMethod,
                    null, wsChart, null);
                if (chartObjects != null)
                {
                    var chartObjectsType = chartObjects.GetType();
                    var count = chartObjectsType.InvokeMember("Count", System.Reflection.BindingFlags.GetProperty,
                        null, chartObjects, null);
                    Log.Debug("Найдено {Count} старых графиков для удаления", count);

                    chartObjectsType.InvokeMember("Delete", System.Reflection.BindingFlags.InvokeMethod,
                        null, chartObjects, null);
                    if (chartObjects != null && Marshal.IsComObject(chartObjects))
                        Marshal.ReleaseComObject(chartObjects);
                    chartObjects = null;
                    Log.Debug("Старые графики удалены");
                }
            }

            // Шаг 12: Создание нового графика
            Log.Debug("Шаг 12: Создание нового ChartObject (left={Left}, top={Top}, width={Width}, height={Height})",
                left, top, width, height);
            chartObjects = wsChartType.InvokeMember("ChartObjects", System.Reflection.BindingFlags.InvokeMethod,
                null, wsChart, null);
            var chartObjectsType2 = chartObjects.GetType();

            chartObj = chartObjectsType2.InvokeMember("Add", System.Reflection.BindingFlags.InvokeMethod,
                null, chartObjects, new object[] { left, top, width, height });
            if (chartObj == null)
            {
                Log.Error("Не удалось создать ChartObject");
                throw new InvalidOperationException("Не удалось создать ChartObject");
            }
            var chartObjType = chartObj.GetType();
            Log.Debug("ChartObject создан");

            // Шаг 13: Получение объекта Chart
            Log.Debug("Шаг 13: Получение Chart из ChartObject");
            chart = chartObjType.InvokeMember("Chart", System.Reflection.BindingFlags.GetProperty,
                null, chartObj, null);
            if (chart == null)
            {
                Log.Error("Не удалось получить Chart из ChartObject");
                throw new InvalidOperationException("Не удалось получить Chart");
            }
            var chartType = chart.GetType();
            Log.Debug("Chart получен");

            // Шаг 14: Настройка источника данных
            Log.Debug("Шаг 14: Установка источника данных (SetSourceData)");
            chartType.InvokeMember("SetSourceData", System.Reflection.BindingFlags.InvokeMethod,
                null, chart, new object[] { dataBodyRange, xlRows });
            Log.Debug("Источник данных установлен");

            // Шаг 15: Установка типа графика
            Log.Debug("Шаг 15: Установка типа графика xlLine={XlLine}", xlLine);
            chartType.InvokeMember("ChartType", System.Reflection.BindingFlags.SetProperty,
                null, chart, new object[] { xlLine });
            Log.Debug("Тип графика установлен");

            // Шаг 16: Сохранение и закрытие
            Log.Debug("Шаг 16: Сохранение workbook");
            wbType.InvokeMember("Save", System.Reflection.BindingFlags.InvokeMethod,
                null, wb, null);
            Log.Debug("Workbook сохранён");

            Log.Debug("Шаг 17: Закрытие workbook");
            wbType.InvokeMember("Close", System.Reflection.BindingFlags.InvokeMethod,
                null, wb, new object[] { false });
            Log.Debug("Workbook закрыт");

            Log.Debug("Шаг 18: Завершение работы Excel");
            excelType.InvokeMember("Quit", System.Reflection.BindingFlags.InvokeMethod,
                null, app, null);
            Log.Information("График успешно создан для файла {FilePath}", filePath);
        }
        catch (Exception ex)
        {
            Log.Error(ex, "Ошибка при создании графика через COM Interop на одном из шагов");
            throw;
        }
        finally
        {
            Log.Debug("Освобождение COM объектов");
            if (chart != null && Marshal.IsComObject(chart)) Marshal.ReleaseComObject(chart);
            if (chartObj != null && Marshal.IsComObject(chartObj)) Marshal.ReleaseComObject(chartObj);
            if (chartObjects != null && Marshal.IsComObject(chartObjects)) Marshal.ReleaseComObject(chartObjects);
            if (dataBodyRange != null && Marshal.IsComObject(dataBodyRange)) Marshal.ReleaseComObject(dataBodyRange);
            if (loRange != null && Marshal.IsComObject(loRange)) Marshal.ReleaseComObject(loRange);
            if (lo != null && Marshal.IsComObject(lo)) Marshal.ReleaseComObject(lo);
            if (listObjects != null && Marshal.IsComObject(listObjects)) Marshal.ReleaseComObject(listObjects);
            if (wsChart != null && Marshal.IsComObject(wsChart)) Marshal.ReleaseComObject(wsChart);
            if (wsData != null && Marshal.IsComObject(wsData)) Marshal.ReleaseComObject(wsData);
            if (worksheets != null && Marshal.IsComObject(worksheets)) Marshal.ReleaseComObject(worksheets);
            if (wb != null && Marshal.IsComObject(wb)) Marshal.ReleaseComObject(wb);
            if (workbooks != null && Marshal.IsComObject(workbooks)) Marshal.ReleaseComObject(workbooks);
            if (app != null && Marshal.IsComObject(app)) Marshal.ReleaseComObject(app);

            System.GC.Collect(); System.GC.WaitForPendingFinalizers();
            System.GC.Collect(); System.GC.WaitForPendingFinalizers();
            Log.Debug("COM объекты освобождены, GC выполнен");
        }
    }
}
