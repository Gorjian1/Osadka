using ClosedXML.Excel;
using Osadka.Core.Units;
using Osadka.Models;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;

namespace Osadka.Services.Data;

public enum CoordUnits { Millimeters, Centimeters, Decimeters, Meters }

public class ExcelImportService
{
    private readonly CoordUnits _coordUnit;

    public ExcelImportService(CoordUnits coordUnit = CoordUnits.Millimeters)
    {
        _coordUnit = coordUnit;
    }

    public ImportResult ImportFromWorkbook(
        IXLWorksheet worksheet,
        List<(int Row, IXLCell Cell)>? objectHeaders,
        List<int>? cycleStarts,
        int selectedObjectIndex,
        int selectedCycleIndex)
    {
        if (worksheet == null)
            throw new ArgumentNullException(nameof(worksheet));

        var result = new ImportResult();

        // 1. Найти заголовки объектов
        if (objectHeaders == null || objectHeaders.Count == 0)
            objectHeaders = FindObjectHeaders(worksheet);

        if (objectHeaders == null || objectHeaders.Count == 0)
            throw new InvalidOperationException("Не удалось найти заголовок с «№ точки» на листе.");

        // 2. Выбрать нужный объект
        var hdrTuple = selectedObjectIndex >= 1 && selectedObjectIndex <= objectHeaders.Count
            ? objectHeaders[selectedObjectIndex - 1]
            : objectHeaders.First();

        int idCol = hdrTuple.Cell.Address.ColumnNumber;
        int subHdrRow = FindSubHeaderRow(worksheet, hdrTuple.Row, idCol);

        // 3. Найти начала циклов
        if (cycleStarts == null || cycleStarts.Count == 0)
        {
            var computed = FindCycleStarts(worksheet, subHdrRow, idCol);
            if (computed.Count == 0)
            {
                int lastRow = worksheet.LastRowUsed().RowNumber();
                for (int r = hdrTuple.Row; r <= Math.Min(hdrTuple.Row + 10, lastRow); r++)
                {
                    bool anyOtm = worksheet.Row(r).Cells()
                        .Any(c => Regex.IsMatch(c.GetString(), @"^\s*Отметка", RegexOptions.IgnoreCase));
                    if (anyOtm)
                    {
                        subHdrRow = r;
                        computed = FindCycleStarts(worksheet, subHdrRow, idCol);
                        if (computed.Count > 0) break;
                    }
                }
            }
            cycleStarts = computed;
        }

        // 4. Прочитать все объекты и циклы
        ReadAllObjects(worksheet, objectHeaders, cycleStarts, result);

        // 5. Определить выбранные объект и цикл
        result.SelectedObjectNumber = selectedObjectIndex >= 1 && selectedObjectIndex <= result.ObjectNumbers.Count
            ? result.ObjectNumbers[selectedObjectIndex - 1]
            : (result.ObjectNumbers.Count > 0 ? result.ObjectNumbers[0] : 1);

        if (result.Objects.TryGetValue(result.SelectedObjectNumber, out var cyclesForObject))
        {
            result.CycleNumbers.AddRange(cyclesForObject.Keys.OrderBy(k => k));
        }

        if (result.CycleNumbers.Count > 0)
        {
            int idx = Math.Clamp(selectedCycleIndex, 1, result.CycleNumbers.Count);
            result.SelectedCycleNumber = result.CycleNumbers[idx - 1];
        }

        return result;
    }

    public static FileStream OpenWorkbookStream(string filePath)
    {
        var shareModes = new[]
        {
            FileShare.ReadWrite | FileShare.Delete,
            FileShare.ReadWrite,
            FileShare.Read
        };

        IOException? lastError = null;
        foreach (var share in shareModes)
        {
            for (int attempt = 0; attempt < 3; attempt++)
            {
                try
                {
                    return new FileStream(filePath, FileMode.Open, FileAccess.Read, share);
                }
                catch (IOException ex)
                {
                    lastError = ex;
                    if (attempt < 2)
                        Thread.Sleep(100);
                }
            }
        }

        throw lastError ?? new IOException($"Не удалось открыть файл '{filePath}'.");
    }

    private List<(int Row, IXLCell Cell)> FindObjectHeaders(IXLWorksheet sheet)
    {
        return sheet.RangeUsed()?
            .Rows()
            .Select(r =>
            {
                var hits = r.Cells()
                    .Where(c => Regex.IsMatch(c.GetString(), @"^\s*№\s*(точки|мар\w*)", RegexOptions.IgnoreCase));
                if (!hits.Any()) return (Row: 0, Cell: (IXLCell?)null);
                var leftMost = hits.OrderBy(c => c.Address.ColumnNumber).First();
                return (Row: r.RowNumber(), Cell: leftMost);
            })
            .Where(t => t.Cell != null && t.Row > 0)
            .ToList()
            ?? new List<(int Row, IXLCell Cell)>();
    }

    private List<int> FindCycleStarts(IXLWorksheet sheet, int subHdrRow, int idColumn)
    {
        return sheet.Row(subHdrRow)
            .Cells()
            .Where(c => c.Address.ColumnNumber != idColumn &&
                       c.GetString().Trim().StartsWith("Отметка", StringComparison.OrdinalIgnoreCase))
            .Select(c => c.Address.ColumnNumber)
            .Distinct()
            .OrderBy(c => c)
            .ToList();
    }

    private int FindSubHeaderRow(IXLWorksheet sheet, int headerRow, int idColumn)
    {
        int lastRow = sheet.LastRowUsed().RowNumber();
        for (int r = headerRow + 1; r <= Math.Min(headerRow + 6, lastRow); r++)
        {
            bool ok = sheet.Row(r).Cells()
                .Any(c => c.Address.ColumnNumber != idColumn &&
                         c.GetString().Trim().StartsWith("Отметка", StringComparison.OrdinalIgnoreCase));
            if (ok) return r;
        }
        return headerRow + 1;
    }

    private void ReadAllObjects(
        IXLWorksheet sheet,
        List<(int Row, IXLCell Cell)> headers,
        List<int> cycleCols,
        ImportResult result)
    {
        if (headers == null || headers.Count == 0) return;

        headers = headers.OrderBy(h => h.Row).ToList();

        for (int objNumber = 1; objNumber <= headers.Count; objNumber++)
        {
            var hdr = headers[objNumber - 1];

            int idColLocal = hdr.Cell.Address.ColumnNumber;
            int subHdrRowLocal = FindSubHeaderRow(sheet, hdr.Row, idColLocal);

            int dataRowFirst = subHdrRowLocal + 1;
            int dataRowLast = objNumber == headers.Count
                ? sheet.LastRowUsed().RowNumber()
                : headers[objNumber].Row - 1;

            var localCycCols = cycleCols != null && cycleCols.Count > 0
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

                    if (mark.HasValue) mark = UnitConverter.ToMm(mark.Value, MapUnit(_coordUnit));
                    if (settl.HasValue) settl = UnitConverter.ToMm(settl.Value, MapUnit(_coordUnit));
                    if (total.HasValue) total = UnitConverter.ToMm(total.Value, MapUnit(_coordUnit));

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
            result.ObjectNumbers.Add(objNumber);
        }

        result.ObjectNumbers.Sort();
    }

    private string BuildCycleHeaderLabel(IXLWorksheet sheet, int startCol, int subHdrRow, int headerRow)
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
                if (!string.IsNullOrWhiteSpace(s) && Regex.IsMatch(s, @"^\s*Цикл\b", RegexOptions.IgnoreCase))
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
                if (!string.IsNullOrWhiteSpace(s) && Regex.IsMatch(s, @"^\s*Цикл\b", RegexOptions.IgnoreCase))
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

    private static Unit MapUnit(CoordUnits units)
    {
        return units switch
        {
            CoordUnits.Millimeters => Unit.Millimeter,
            CoordUnits.Centimeters => Unit.Centimeter,
            CoordUnits.Decimeters => Unit.Decimeter,
            CoordUnits.Meters => Unit.Meter,
            _ => Unit.Millimeter
        };
    }
}

public class ImportResult
{
    public Dictionary<int, Dictionary<int, List<MeasurementRow>>> Objects { get; } = new();
    public List<int> ObjectNumbers { get; } = new();
    public List<int> CycleNumbers { get; } = new();
    public Dictionary<int, string> CycleHeaders { get; } = new();
    public int SelectedObjectNumber { get; set; } = 1;
    public int SelectedCycleNumber { get; set; } = 1;
}
