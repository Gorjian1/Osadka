using Osadka.Core.Units;
using Osadka.Models;
using Osadka.Services.Abstractions;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;

namespace Osadka.Services.Implementation;

/// <summary>
/// Реализация сервиса парсинга данных из буфера обмена
/// </summary>
public class ClipboardParserService : IClipboardParserService
{
    public IClipboardParserService.ParseResult Parse(
        string clipboardText,
        int cycleNumber,
        IReadOnlyList<string> existingIds,
        Unit coordUnit)
    {
        var result = new IClipboardParserService.ParseResult();

        if (string.IsNullOrWhiteSpace(clipboardText))
            return result;

        var lines = clipboardText.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
        if (lines.Length == 0)
            return result;

        var first = lines[0].Split('\t');
        int cols = first.Length;

        // 1 колонка: ID точек
        if (cols == 1)
        {
            result.Type = IClipboardParserService.ParseResult.DataType.Ids;
            foreach (var ln in lines)
            {
                result.Ids.Add(ln.Trim());
            }
            return result;
        }

        // 2 колонки: X, Y координаты
        if (cols == 2)
        {
            result.Type = IClipboardParserService.ParseResult.DataType.Coordinates;

            foreach (var ln in lines)
            {
                var arr = ln.Split('\t');
                if (arr.Length < 2) continue;

                if (!double.TryParse(arr[0].Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out double x) ||
                    !double.TryParse(arr[1].Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out double y))
                    continue;

                result.Coordinates.Add(new CoordRow
                {
                    X = UnitConverter.ToMm(x, coordUnit),
                    Y = UnitConverter.ToMm(y, coordUnit)
                });
            }
            return result;
        }

        // 3 колонки: Отметка, Осадка, Общая (ID берутся из existingIds)
        if (cols == 3)
        {
            result.Type = IClipboardParserService.ParseResult.DataType.Measurements3;
            int row = 0;

            foreach (var ln in lines)
            {
                var arr = ln.Split('\t');
                if (arr.Length < 3) continue;
                if (LooksLikeHeader(arr)) continue;

                var (markVal, markRaw) = TryParse(arr[0]);
                var (settlVal, settlRaw) = TryParse(arr[1]);
                var (totalVal, totalRaw) = TryParse(arr[2]);

                string id = (row < existingIds.Count) ? existingIds[row] : (row + 1).ToString();

                result.Measurements.Add(new MeasurementRow
                {
                    Id = id,
                    Mark = markVal,
                    Settl = settlVal,
                    Total = totalVal,
                    MarkRaw = markRaw,
                    SettlRaw = settlRaw,
                    TotalRaw = totalRaw,
                    Cycle = cycleNumber
                });
                row++;
            }
            return result;
        }

        // 4 колонки: Отметка, Осадка, Общая, ID
        if (cols == 4)
        {
            result.Type = IClipboardParserService.ParseResult.DataType.Measurements4;

            foreach (var ln in lines)
            {
                var arr = ln.Split('\t');
                if (arr.Length < 4) continue;
                if (LooksLikeHeader(arr)) continue;

                string markRaw = arr[0];
                string settlRaw = arr[1];
                string totalRaw = arr[2];
                string id = arr[3].Trim();
                if (string.IsNullOrEmpty(id)) continue;

                var (markVal, _) = TryParse(markRaw);
                var (settlVal, _) = TryParse(settlRaw);
                var (totalVal, _) = TryParse(totalRaw);

                result.Measurements.Add(new MeasurementRow
                {
                    Id = id,
                    Mark = markVal,
                    Settl = settlVal,
                    Total = totalVal,
                    MarkRaw = markRaw,
                    SettlRaw = settlRaw,
                    TotalRaw = totalRaw,
                    Cycle = cycleNumber
                });
            }
            return result;
        }

        // Неподдерживаемый формат
        return result;
    }

    private static (double? val, string raw) TryParse(string txt)
    {
        txt = txt.Trim();
        if (Regex.IsMatch(txt, @"\bнов", RegexOptions.IgnoreCase))
            return (0, txt);

        if (double.TryParse(txt.Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out double v))
            return (v, txt);

        return (null, txt);
    }

    private static bool LooksLikeHeader(string[] cells)
    {
        var joined = string.Join(" ", cells).ToLowerInvariant();
        if (joined.Contains("отмет") || joined.Contains("осад") || joined.Contains("суммар") ||
            joined.Contains("№") || joined.Contains("марка") || joined.Contains("cycle") || joined.Contains("id"))
            return true;

        int nonNumeric = 0;
        for (int i = 0; i < cells.Length; i++)
        {
            var t = (cells[i] ?? string.Empty).Trim();
            if (Regex.IsMatch(t, @"\bнов", RegexOptions.IgnoreCase))
                continue;

            if (!double.TryParse(t.Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out _))
                nonNumeric++;
        }
        return nonNumeric >= Math.Max(2, cells.Length - 1);
    }
}
