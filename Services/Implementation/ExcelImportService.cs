using ClosedXML.Excel;
using Osadka.Core.Units;
using Osadka.Models;
using Osadka.Services.Abstractions;
using Osadka.ViewModels;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows;

namespace Osadka.Services.Implementation;

/// <summary>
/// Реализация сервиса импорта данных из Excel
/// </summary>
public class ExcelImportService : IExcelImportService
{
    public IExcelImportService.ImportResult? ImportFromExcel(string filePath, RawDataViewModel.CoordUnits coordUnit)
    {
        if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
            return null;

        try
        {
            using var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
            using var wb = new XLWorkbook(stream);

            var dlg = new Osadka.Views.ImportSelectionWindow(wb)
            {
                Owner = Application.Current?.MainWindow
            };
            if (dlg.ShowDialog() != true)
                return null;

            IXLWorksheet ws = dlg.SelectedWorksheet?.Sheet
                ?? throw new InvalidOperationException("Не выбран лист Excel.");

            var objHeaders = dlg.ObjectHeaders;
            var cycleStarts = dlg.CycleStarts;
            int objIdx = dlg.SelectedObjectIndex;   // 1-based
            int cycleIdx = dlg.SelectedCycleIndex;  // 1-based

            if (objHeaders == null || objHeaders.Count == 0)
                objHeaders = FindObjectHeaders(ws);
            if (objHeaders == null || objHeaders.Count == 0)
                throw new InvalidOperationException("Не удалось найти заголовок с «№ точки» на листе.");

            var hdrTuple = objIdx >= 1 && objIdx <= objHeaders.Count
                ? objHeaders[objIdx - 1]
                : objHeaders.First();
            int idCol = hdrTuple.Cell.Address.ColumnNumber;
            int subHdrRow = FindSubHeaderRow(ws, hdrTuple.Row, idCol);

            if (cycleStarts == null || cycleStarts.Count == 0)
            {
                var computed = FindCycleStarts(ws, subHdrRow, idCol);
                if (computed.Count == 0)
                {
                    int lastRow = ws.LastRowUsed().RowNumber();
                    for (int r = hdrTuple.Row; r <= Math.Min(hdrTuple.Row + 10, lastRow); r++)
                    {
                        bool anyOtm = ws.Row(r).Cells().Any(c =>
                            Regex.IsMatch(c.GetString(), @"^\s*Отметка", RegexOptions.IgnoreCase));
                        if (anyOtm)
                        {
                            subHdrRow = r;
                            computed = FindCycleStarts(ws, subHdrRow, idCol);
                            if (computed.Count > 0) break;
                        }
                    }
                }
                cycleStarts = computed;
            }

            var result = new IExcelImportService.ImportResult();
            ReadAllObjects(ws, objHeaders, cycleStarts, coordUnit, result);

            // Определяем рекомендуемые номера объекта и цикла
            var objectNumbers = result.Objects.Keys.OrderBy(k => k).ToList();
            if (objectNumbers.Count > 0)
            {
                result.SuggestedObjectNumber = (objIdx >= 1 && objIdx <= objectNumbers.Count)
                    ? objectNumbers[objIdx - 1]
                    : objectNumbers[0];

                if (result.Objects.TryGetValue(result.SuggestedObjectNumber, out var cycles))
                {
                    var cycleNumbers = cycles.Keys.OrderBy(k => k).ToList();
                    if (cycleNumbers.Count > 0)
                    {
                        int idx = Math.Clamp(cycleIdx, 1, cycleNumbers.Count);
                        result.SuggestedCycleNumber = cycleNumbers[idx - 1];
                    }
                }
            }

            return result;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Ошибка при импорте Excel: {ex.Message}", ex);
        }
    }

    private static List<(int Row, IXLCell Cell)> FindObjectHeaders(IXLWorksheet sheet)
        => sheet.RangeUsed()?
               .Rows()
               .Select(r =>
               {
                   var hits = r.Cells().Where(c =>
                       Regex.IsMatch(c.GetString(), @"^\s*№\s*(точки|мар\w*)", RegexOptions.IgnoreCase));
                   if (!hits.Any())
                       return (Row: 0, Cell: (IXLCell?)null);
                   var leftMost = hits.OrderBy(c => c.Address.ColumnNumber).First();
                   return (Row: r.RowNumber(), Cell: leftMost);
               })
               .Where(t => t.Cell != null && t.Row > 0)
               .ToList()
           ?? new List<(int Row, IXLCell Cell)>();

    private static List<int> FindCycleStarts(IXLWorksheet sheet, int subHdrRow, int idColumn)
        => sheet.Row(subHdrRow)
                .Cells()
                .Where(c => c.Address.ColumnNumber != idColumn &&
                           c.GetString().Trim().StartsWith("Отметка", StringComparison.OrdinalIgnoreCase))
                .Select(c => c.Address.ColumnNumber)
                .Distinct()
                .OrderBy(c => c)
                .ToList();

    private static int FindSubHeaderRow(IXLWorksheet s, int headerRow, int idColumn)
    {
        int lastRow = s.LastRowUsed().RowNumber();
        for (int r = headerRow + 1; r <= Math.Min(headerRow + 6, lastRow); r++)
        {
            bool ok = s.Row(r).Cells().Any(c =>
                c.Address.ColumnNumber != idColumn &&
                c.GetString().Trim().StartsWith("Отметка", StringComparison.OrdinalIgnoreCase));
            if (ok) return r;
        }
        return headerRow + 1;
    }

    private static void ReadAllObjects(
        IXLWorksheet sheet,
        List<(int Row, IXLCell Cell)> headers,
        List<int> cycleCols,
        RawDataViewModel.CoordUnits coordUnit,
        IExcelImportService.ImportResult result)
    {
        if (headers == null || headers.Count == 0)
            return;

        headers = headers.OrderBy(h => h.Row).ToList();

        for (int objNumber = 1; objNumber <= headers.Count; objNumber++)
        {
            var hdr = headers[objNumber - 1];

            int idColLocal = hdr.Cell.Address.ColumnNumber;
            int subHdrRowLocal = FindSubHeaderRow(sheet, hdr.Row, idColLocal);

            int dataRowFirst = subHdrRowLocal + 1;
            int dataRowLast = (objNumber == headers.Count
                ? sheet.LastRowUsed().RowNumber()
                : headers[objNumber].Row - 1);

            var localCycCols = (cycleCols != null && cycleCols.Count > 0)
                ? cycleCols
                : FindCycleStarts(sheet, subHdrRowLocal, idColLocal);

            var cyclesDict = new Dictionary<int, List<MeasurementRow>>();

            foreach (var (cycIdx, startCol) in localCycCols.Select((c, i) => (i + 1, c)))
            {
                string cycLabel = BuildCycleHeaderLabel(sheet, startCol, subHdrRowLocal, hdr.Row);
                if (!string.IsNullOrWhiteSpace(cycLabel))
                    result.CycleHeaders[cycIdx] = cycLabel;

                var rows = new List<MeasurementRow>();
                int blanksInARow = 0;

                for (int r = dataRowFirst; r <= dataRowLast; r++)
                {
                    string idText = sheet.Cell(r, idColLocal).GetString().Trim();
                    if (Regex.IsMatch(idText, @"^\s*№\s*(точки|мар\w*)", RegexOptions.IgnoreCase))
                        break;

                    if (string.IsNullOrEmpty(idText))
                    {
                        blanksInARow++;
                        if (blanksInARow >= 3) break;
                        continue;
                    }
                    blanksInARow = 0;

                    var (mark, markRaw) = ParseCell(sheet.Cell(r, startCol));
                    var (settl, settlRaw) = ParseCell(sheet.Cell(r, startCol + 1));
                    var (total, totalRaw) = ParseCell(sheet.Cell(r, startCol + 2));

                    // Конвертируем в мм
                    var unit = MapCoordUnit(coordUnit);
                    if (mark.HasValue) mark = UnitConverter.ToMm(mark.Value, unit);
                    if (settl.HasValue) settl = UnitConverter.ToMm(settl.Value, unit);
                    if (total.HasValue) total = UnitConverter.ToMm(total.Value, unit);

                    if (mark is null && settl is null && total is null &&
                        string.IsNullOrWhiteSpace(markRaw) &&
                        string.IsNullOrWhiteSpace(settlRaw) &&
                        string.IsNullOrWhiteSpace(totalRaw))
                    {
                        continue;
                    }

                    if (settl.HasValue) settl = Math.Round(settl.Value, 1);
                    if (total.HasValue) total = Math.Round(total.Value, 1);

                    rows.Add(new MeasurementRow
                    {
                        Id = idText,
                        Mark = mark,
                        Settl = settl,
                        Total = total,
                        MarkRaw = markRaw,
                        SettlRaw = settlRaw,
                        TotalRaw = totalRaw
                    });
                }

                cyclesDict[cycIdx] = rows;
            }

            result.Objects[objNumber] = cyclesDict;
        }
    }

    private static string BuildCycleHeaderLabel(IXLWorksheet sheet, int startCol, int subHdrRow, int headerRow)
    {
        string Read(IXLCell cell)
        {
            var s = cell.GetString();
            if (!string.IsNullOrWhiteSpace(s)) return s;
            var mr = cell.MergedRange();
            return mr != null ? mr.FirstCell().GetString() : s;
        }

        int r1 = Math.Max(1, headerRow - 2);
        int r2 = subHdrRow + 1;

        // 1) Ищем только внутри текущей тройки (Отметка/Осадка/Общая)
        for (int r = r1; r <= r2; r++)
        {
            for (int c = startCol; c <= startCol + 2; c++)
            {
                var s = Read(sheet.Cell(r, c));
                if (!string.IsNullOrWhiteSpace(s) &&
                    Regex.IsMatch(s, @"^\s*Цикл\b", RegexOptions.IgnoreCase))
                    return s.Trim();
            }
        }

        // 2) Фолбэк — центр-сначала (0,+1,-1,+2,-2,...)
        int[] offs = new[] { 0, +1, -1, +2, -2, +3, -3 };
        for (int r = r1; r <= r2; r++)
        {
            foreach (var dc in offs)
            {
                int c = startCol + dc;
                if (c <= 0) continue;
                var s = Read(sheet.Cell(r, c));
                if (!string.IsNullOrWhiteSpace(s) &&
                    Regex.IsMatch(s, @"^\s*Цикл\b", RegexOptions.IgnoreCase))
                    return s.Trim();
            }
        }

        return string.Empty;
    }

    private static (double? val, string raw) ParseCell(IXLCell cell)
    {
        string txt = cell.GetString().Trim();
        if (Regex.IsMatch(txt, @"\bнов", RegexOptions.IgnoreCase))
            return (0, txt);

        if (cell.DataType == XLDataType.Number)
            return (cell.GetDouble(), txt);

        if (double.TryParse(txt.Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out double v))
            return (v, txt);

        return (null, txt);
    }

    private static Unit MapCoordUnit(RawDataViewModel.CoordUnits u) => u switch
    {
        RawDataViewModel.CoordUnits.Millimeters => Unit.Millimeter,
        RawDataViewModel.CoordUnits.Centimeters => Unit.Centimeter,
        RawDataViewModel.CoordUnits.Decimeters => Unit.Decimeter,
        _ => Unit.Meter
    };
}
