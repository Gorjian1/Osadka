using Osadka.Models;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace Osadka.Services
{
    public record Extremum(double Value, IReadOnlyList<string> Ids);

    public record GeneralReportData(
        Extremum MaxTotal, Extremum MinTotal, double? AvgTotal,
        Extremum MaxSettl, Extremum MinSettl, double? AvgSettl,
        IReadOnlyList<string> NoAccessIds,
        IReadOnlyList<string> NewIds,
        IReadOnlyList<string> DestroyedIds,
        IReadOnlyList<string> ExceedTotalSpIds,
        IReadOnlyList<string> ExceedTotalCalcIds);

    public sealed class GeneralReportService
    {
        public GeneralReportData Build(
            IEnumerable<MeasurementRow> rows,
            double limitSp,
            double limitCalc)
        {
            var total = rows.Where(r => r.Total is { } v && !double.IsNaN(v))
                            .ToList();
            var settl = rows.Where(r => r.Settl is { } v && !double.IsNaN(v))
                            .ToList();

            Extremum GetMax(IEnumerable<MeasurementRow> src, Func<MeasurementRow, double?> sel)
            {
                if (!src.Any())
                    return new Extremum(double.NaN, Array.Empty<string>());

                double max = Math.Round(src.Max(r => sel(r)!.Value), 4);
                var ids = src.Where(r => Math.Abs(sel(r)!.Value - max) < 1e-6)
                             .Select(r => r.Id)
                             .ToList();
                return new Extremum(max, ids);
            }

            Extremum GetMin(IEnumerable<MeasurementRow> src, Func<MeasurementRow, double?> sel)
            {
                if (!src.Any())
                    return new Extremum(double.NaN, Array.Empty<string>());

                double min = Math.Round(src.Min(r => sel(r)!.Value), 4);
                var ids = src.Where(r => Math.Abs(sel(r)!.Value - min) < 1e-6)
                             .Select(r => r.Id)
                             .ToList();
                return new Extremum(min, ids);
            }

            var maxTotal = GetMax(total, r => r.Total);
            var minTotal = GetMin(total, r => r.Total);
            double? avgTotal = total.Any()
                ? Math.Round(total.Average(r => r.Total!.Value), 4)
                : null;

            var maxSettl = GetMax(settl, r => r.Settl);
            var minSettl = GetMin(settl, r => r.Settl);
            double? avgSettl = settl.Any()
                ? Math.Round(settl.Average(r => r.Settl!.Value), 4)
                : null;
            var noAccess = rows
                .Where(r => r.MarkRaw.Contains("нет доступ", StringComparison.OrdinalIgnoreCase))
                .Select(r => r.Id)
                .ToList();

            var @new = rows
                .Where(r => r.SettlRaw.Contains("нов", StringComparison.OrdinalIgnoreCase))
                .Select(r => r.Id)
                .ToList();

            var destroyed = rows
                .Where(r => r.MarkRaw.Contains("унич", StringComparison.OrdinalIgnoreCase))
                .Select(r => r.Id)
                .ToList();

            bool Exceeded(double x, double lim) => Math.Abs(x) > lim;

            var exceedSp = total
                .Where(r => Exceeded(r.Total!.Value, limitSp))
                .Select(r => r.Id)
                .ToList();

            var exceedCalc = total
                .Where(r => Exceeded(r.Total!.Value, limitCalc))
                .Select(r => r.Id)
                .ToList();

            return new GeneralReportData(
                maxTotal, minTotal, avgTotal,
                maxSettl, minSettl, avgSettl,
                noAccess, @new, destroyed,
                exceedSp, exceedCalc);
        }
    }
}
