using ClosedXML.Excel;
using Osadka.Models;
using Osadka.Services.Parsing;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;

namespace Osadka.Services.Reports;

public class ExcelExportOptions
{
    public bool IncludeGeneral { get; set; }
    public bool IncludeRelative { get; set; }
    public bool IncludeGraphs { get; set; }
    public string TemplatePath { get; set; } = string.Empty;
    public string OutputPath { get; set; } = string.Empty;
}

public class ExcelExportData
{
    public Dictionary<string, string> PlaceholderMap { get; set; } = new();
    public HashSet<string> DisabledTags { get; set; } = new();
    public List<RelativeSettlementRow> RelativeRows { get; set; } = new();
    public Dictionary<int, List<MeasurementRow>> ActiveCycles { get; set; } = new();
    public Dictionary<int, string> CycleHeaders { get; set; } = new();
    public List<DynamicsSeries> DynamicsData { get; set; } = new();
}

public class RelativeSettlementRow
{
    public string Id1 { get; set; } = string.Empty;
    public string Id2 { get; set; } = string.Empty;
    public double Distance { get; set; }
    public double DeltaTotal { get; set; }
    public double Ratio { get; set; }
}

public class DynamicsSeries
{
    public string Id { get; set; } = string.Empty;
    public List<DynamicsPoint> Points { get; set; } = new();
}

public class DynamicsPoint
{
    public int Cycle { get; set; }
    public double Mark { get; set; }
}

public class ExcelExportService
{
    public void ExportToExcel(ExcelExportOptions options, ExcelExportData data)
    {
        if (options == null)
            throw new ArgumentNullException(nameof(options));
        if (data == null)
            throw new ArgumentNullException(nameof(data));

        using (var wb = new XLWorkbook(options.OutputPath))
        {
            var generalWs = wb.Worksheets.First(); // титульный/общий

            if (options.IncludeGeneral)
            {
                ProcessGeneralSheet(generalWs, data.PlaceholderMap, data.DisabledTags);
            }
            else
            {
                generalWs.Delete();
            }

            if (options.IncludeRelative)
                AddRelativeSheet(wb, data.RelativeRows);

            if (options.IncludeGraphs)
            {
                AddDynamicsSheet(wb, data.ActiveCycles, data.CycleHeaders, data.DynamicsData);
            }
            else
            {
                var dynWs = wb.Worksheets.FirstOrDefault(
                    ws => string.Equals(ws.Name, "Графики динамики", StringComparison.OrdinalIgnoreCase));
                dynWs?.Delete();
            }

            wb.Save();
        }
    }

    public void BuildExcelChart(
        string filePath,
        string dataSheetName = "Графики динамики",
        string tableName = "DynTable",
        string chartSheetName = "Графики динамики",
        int left = 40, int top = 200, int width = 920, int height = 440,
        bool deleteOldCharts = true)
    {
        const int xlRows = 1;
        const int xlLine = 4;

        object? app = null, wb = null, wsData = null, wsChart = null;
        object? workbooks = null, worksheets = null;
        object? listObjects = null, lo = null, loRange = null, dataBodyRange = null;
        object? chartObjects = null, chartObj = null, chart = null;

        try
        {
            var excelType = Type.GetTypeFromProgID("Excel.Application", throwOnError: false);
            if (excelType == null)
                throw new InvalidOperationException("На этом ПК не установлен Microsoft Excel.");

            app = Activator.CreateInstance(excelType);
            excelType.InvokeMember("Visible",
                System.Reflection.BindingFlags.SetProperty, null, app, new object[] { false });
            excelType.InvokeMember("DisplayAlerts",
                System.Reflection.BindingFlags.SetProperty, null, app, new object[] { false });

            // Workbooks.Open(filePath)
            workbooks = excelType.InvokeMember("Workbooks",
                System.Reflection.BindingFlags.GetProperty, null, app, null);
            wb = workbooks!.GetType().InvokeMember("Open",
                System.Reflection.BindingFlags.InvokeMethod, null, workbooks, new object[] { filePath });

            // Worksheets[dataSheetName]
            worksheets = wb!.GetType().InvokeMember("Worksheets",
                System.Reflection.BindingFlags.GetProperty, null, wb, null);
            try
            {
                wsData = worksheets!.GetType().InvokeMember("Item",
                    System.Reflection.BindingFlags.GetProperty, null, worksheets, new object[] { dataSheetName });
            }
            catch
            {
                throw new InvalidOperationException($"Не найден лист данных '{dataSheetName}'.");
            }

            // Таблица ListObjects[tableName]
            listObjects = wsData!.GetType().InvokeMember("ListObjects",
                System.Reflection.BindingFlags.GetProperty, null, wsData, null);
            try
            {
                lo = listObjects!.GetType().InvokeMember("Item",
                    System.Reflection.BindingFlags.GetProperty, null, listObjects, new object[] { tableName });
            }
            catch
            {
                throw new InvalidOperationException($"Не найдена таблица '{tableName}' на листе '{dataSheetName}'.");
            }

            dataBodyRange = lo!.GetType().InvokeMember("DataBodyRange",
                System.Reflection.BindingFlags.GetProperty, null, lo, null);
            if (dataBodyRange == null)
                throw new InvalidOperationException("В таблице нет строк данных.");
            loRange = lo.GetType().InvokeMember("Range",
                System.Reflection.BindingFlags.GetProperty, null, lo, null);

            try
            {
                wsChart = worksheets!.GetType().InvokeMember("Item",
                    System.Reflection.BindingFlags.GetProperty, null, worksheets, new object[] { chartSheetName });
            }
            catch
            {
                wsChart = worksheets!.GetType().InvokeMember("Add",
                    System.Reflection.BindingFlags.InvokeMethod, null, worksheets, null);
                wsChart!.GetType().InvokeMember("Name",
                    System.Reflection.BindingFlags.SetProperty, null, wsChart, new object[] { chartSheetName });
            }

            object missing = Type.Missing;
            try
            {
                chartObjects = wsChart!.GetType().InvokeMember("ChartObjects",
                    System.Reflection.BindingFlags.InvokeMethod, null, wsChart, new object[] { missing });
            }
            catch
            {
                chartObjects = wsChart!.GetType().InvokeMember("ChartObjects",
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
                        int cnt = cntObj is int i ? i : Convert.ToInt32(cntObj);
                        for (int j = cnt; j >= 1; j--)
                        {
                            var co = chartObjects.GetType().InvokeMember("Item",
                                System.Reflection.BindingFlags.GetProperty, null, chartObjects, new object[] { j });
                            co!.GetType().InvokeMember("Delete",
                                System.Reflection.BindingFlags.InvokeMethod, null, co, null);
                            Marshal.FinalReleaseComObject(co);
                        }
                    }
                    catch { }
                }

                try
                {
                    chartObjects = wsChart!.GetType().InvokeMember("ChartObjects",
                        System.Reflection.BindingFlags.InvokeMethod, null, wsChart, new object[] { missing });
                }
                catch
                {
                    chartObjects = wsChart!.GetType().InvokeMember("ChartObjects",
                        System.Reflection.BindingFlags.InvokeMethod, null, wsChart, null);
                }
            }

            chartObj = chartObjects!.GetType().InvokeMember("Add",
                System.Reflection.BindingFlags.InvokeMethod, null, chartObjects,
                new object[] { left, top, width, height });
            chart = chartObj!.GetType().InvokeMember("Chart",
                System.Reflection.BindingFlags.GetProperty, null, chartObj, null);

            chart!.GetType().InvokeMember("SetSourceData",
                System.Reflection.BindingFlags.InvokeMethod, null, chart, new object[] { loRange!, xlRows });
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

            void rel(object? o) { if (o != null) Marshal.FinalReleaseComObject(o); }
            rel(chart); rel(chartObj); rel(chartObjects);
            rel(wsChart); rel(loRange); rel(dataBodyRange); rel(lo); rel(listObjects);
            rel(wsData); rel(worksheets); rel(wb); rel(workbooks); rel(app);

            GC.Collect(); GC.WaitForPendingFinalizers();
            GC.Collect(); GC.WaitForPendingFinalizers();
        }
    }

    private void ProcessGeneralSheet(IXLWorksheet sheet, Dictionary<string, string> placeholders, HashSet<string> disabledTags)
    {
        // Снимок текстовых ячеек, чтобы удалять строки уже после прохода
        var textCells = sheet.CellsUsed(c => c.DataType == XLDataType.Text).ToList();
        var rowsToDelete = new HashSet<int>();

        foreach (var cell in textCells)
        {
            string t = cell.GetString().Trim();
            if (!t.StartsWith("/")) continue;

            if (disabledTags.Contains(t))
            {
                rowsToDelete.Add(cell.Address.RowNumber);
                continue;
            }

            if (placeholders.TryGetValue(t, out var val))
                cell.Value = val;
        }

        // Удаляем строки снизу вверх
        foreach (var r in rowsToDelete.OrderByDescending(x => x))
            sheet.Row(r).Delete();
    }

    private void AddRelativeSheet(XLWorkbook wb, List<RelativeSettlementRow> rows)
    {
        // Получаем/создаём лист
        var ws = wb.Worksheets.FirstOrDefault(s =>
                     s.Name.Equals("Относительная разность", StringComparison.OrdinalIgnoreCase))
                 ?? wb.AddWorksheet("Относительная разность");
        ws.Clear();

        // Заголовки
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
                cell.Style.NumberFormat.Format = "@"; // принудительно "Текст"
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
        foreach (var row in rows)
        {
            ws.Cell(r, 1).Value = row.Id1;
            ws.Cell(r, 2).Value = row.Id2;

            SetNumberOrDash(ws.Cell(r, 3), row.Distance, "0.0");
            SetNumberOrDash(ws.Cell(r, 4), row.DeltaTotal, "0.0");
            SetNumberOrDash(ws.Cell(r, 5), row.Ratio, "0.000000");

            r++;
        }

        // Оформление
        ws.Range(1, 1, 1, 5).Style.Font.Bold = true;
        ws.Columns(1, 5).AdjustToContents();
        ws.SheetView.FreezeRows(1);
    }

    private void AddDynamicsSheet(
        XLWorkbook wb,
        Dictionary<int, List<MeasurementRow>> activeCycles,
        Dictionary<int, string> cycleHeaders,
        List<DynamicsSeries> seriesData)
    {
        const string sheetName = "Графики динамики";
        const string tableName = "DynTable";

        var ws = wb.Worksheets.FirstOrDefault(s =>
                     s.Name.Equals(sheetName, StringComparison.OrdinalIgnoreCase))
                 ?? wb.AddWorksheet(sheetName);

        var cycles = activeCycles.Keys.OrderBy(c => c).ToList();

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
            if (cycleHeaders.TryGetValue(cyc, out var rawLabel))
                headerText = CycleLabelParsing.ExtractDateTail(rawLabel) ?? rawLabel;
            else
                headerText = $"Cycle {cyc}";

            ws.Cell(1, i + 2).Value = Unique(headerText.Trim());
        }

        var colByCycle = cycles
            .Select((cycle, idx) => new { cycle, col = idx + 2 })
            .ToDictionary(x => x.cycle, x => x.col);

        int r = 2;
        foreach (var ser in seriesData)
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
        }
    }
}
