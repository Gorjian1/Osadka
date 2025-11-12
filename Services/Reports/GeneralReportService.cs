using Osadka.Models;
using System;
using System.Linq;
using System.Collections.Generic;

namespace Osadka.Services.Reports
{
    public record Extremum(double Value, IReadOnlyList<string> Ids);
    public record GeneralReportData(
        Extremum MaxTotal, Extremum MinTotal, double? AvgTotal,
        Extremum MaxSettl, Extremum MinSettl, double? AvgSettl,
        IReadOnlyList<string> NoAccessIds,
        IReadOnlyList<string> NewIds,
        IReadOnlyList<string> DestroyedIds,
        IReadOnlyList<string> ExceedTotalSpIds,
        IReadOnlyList<string> ExceedTotalCalcIds,
        string TotalExtrema,
        string SettlExtrema,
        string TotalExtremaIds,
        string SettlExtremaIds
    );

    public sealed class GeneralReportService
    {
        public GeneralReportData Build(
            IEnumerable<MeasurementRow> rows,
            double limitSp,
            double limitCalc)
        {
            var total = rows.Where(r => r.Total is { } v && !double.IsNaN(v)).ToList();
            var settl = rows.Where(r => r.Settl is { } v && !double.IsNaN(v)).ToList();

            static Extremum GetMax(IEnumerable<MeasurementRow> src, Func<MeasurementRow, double?> sel)
            {
                if (!src.Any()) return new Extremum(double.NaN, Array.Empty<string>());
                var max = src.Max(r => sel(r)!.Value);
                var ids = src.Where(r => Math.Abs(sel(r)!.Value - max) < 1e-9).Select(r => r.Id).ToList();
                return new Extremum(Math.Round(max, 4), ids);
            }

            static Extremum GetMin(IEnumerable<MeasurementRow> src, Func<MeasurementRow, double?> sel)
            {
                if (!src.Any()) return new Extremum(double.NaN, Array.Empty<string>());
                var min = src.Min(r => sel(r)!.Value);
                var ids = src.Where(r => Math.Abs(sel(r)!.Value - min) < 1e-9).Select(r => r.Id).ToList();
                return new Extremum(Math.Round(min, 4), ids);
            }

            static string FormatSigned(double v, int decimals = 2)
                => (v >= 0 ? "+" : "") + v.ToString($"F{decimals}");

            static string JoinIdsOrDash(IEnumerable<string> ids)
            {
                var list = ids?.Where(s => !string.IsNullOrWhiteSpace(s)).ToList() ?? new();
                return list.Count > 0 ? string.Join(", ", list) : "-";
            }

            static string BuildExtremumValueString(IEnumerable<MeasurementRow> src, Func<MeasurementRow, double?> sel, int decimals = 2)
            {
                var vals = src.Select(sel).Where(v => v.HasValue && !double.IsNaN(v!.Value))
                              .Select(v => Math.Round(v!.Value, decimals)).ToList();
                if (vals.Count == 0) return "-";

                var negs = vals.Where(v => v < 0).ToList();
                var poss = vals.Where(v => v > 0).ToList();

                if (negs.Count > 0 && poss.Count > 0)
                    return $"{FormatSigned(negs.Min(), decimals)}/{FormatSigned(poss.Max(), decimals)}";
                if (negs.Count > 0)
                    return $"{FormatSigned(negs.Min(), decimals)}";
                if (poss.Count > 0)
                    return $"{FormatSigned(poss.Max(), decimals)}";
                return "-";
            }

            static string BuildExtremumIdsString(IEnumerable<MeasurementRow> src, Func<MeasurementRow, double?> sel, int decimals = 2)
            {
                var data = src.Select(r => (r.Id, V: sel(r)))
                              .Where(t => t.V.HasValue && !double.IsNaN(t.V!.Value))
                              .Select(t => (t.Id, V: Math.Round(t.V!.Value, decimals)))
                              .ToList();

                if (data.Count == 0) return "-";

                var negs = data.Where(x => x.V < 0).ToList();
                var poss = data.Where(x => x.V > 0).ToList();

                if (negs.Count > 0 && poss.Count > 0)
                {
                    var minNeg = negs.Min(x => x.V);
                    var maxPos = poss.Max(x => x.V);
                    var negIds = negs.Where(x => Math.Abs(x.V - minNeg) < 1e-9).Select(x => x.Id);
                    var posIds = poss.Where(x => Math.Abs(x.V - maxPos) < 1e-9).Select(x => x.Id);
                    return $"{JoinIdsOrDash(negIds)} / {JoinIdsOrDash(posIds)}";
                }
                if (negs.Count > 0)
                {
                    var minNeg = negs.Min(x => x.V);
                    var negIds = negs.Where(x => Math.Abs(x.V - minNeg) < 1e-9).Select(x => x.Id);
                    return JoinIdsOrDash(negIds);
                }
                if (poss.Count > 0)
                {
                    var maxPos = poss.Max(x => x.V);
                    var posIds = poss.Where(x => Math.Abs(x.V - maxPos) < 1e-9).Select(x => x.Id);
                    return JoinIdsOrDash(posIds);
                }
                return "-";
            }

            var maxTotal = GetMax(total, r => r.Total);
            var minTotal = GetMin(total, r => r.Total);
            double? avgTotal = total.Any() ? Math.Round(total.Average(r => r.Total!.Value), 4) : (double?)null;

            var maxSettl = GetMax(settl, r => r.Settl);
            var minSettl = GetMin(settl, r => r.Settl);
            double? avgSettl = settl.Any() ? Math.Round(settl.Average(r => r.Settl!.Value), 4) : (double?)null;

            var noAccess = rows.Where(r => r.MarkRaw.Contains("нет доступ", StringComparison.OrdinalIgnoreCase)).Select(r => r.Id).ToList();
            var @new = rows.Where(r => r.SettlRaw.Contains("нов", StringComparison.OrdinalIgnoreCase)).Select(r => r.Id).ToList();
            var destroyed = rows.Where(r => r.MarkRaw.Contains("унич", StringComparison.OrdinalIgnoreCase)).Select(r => r.Id).ToList();

            static bool Exceeded(double x, double lim)
            {
                // Лимит выключен: пусто/NaN/∞/≤0 — не считаем превышения
                if (!double.IsFinite(lim) || lim <= 0) return false;

                return x < -Math.Abs(lim);
            }
            var exceedSp = total.Where(r => Exceeded(r.Total!.Value, limitSp)).Select(r => r.Id).ToList();
            var exceedCalc = total.Where(r => Exceeded(r.Total!.Value, limitCalc)).Select(r => r.Id).ToList();

            var totalExtrema = BuildExtremumValueString(total, r => r.Total);
            var settlExtrema = BuildExtremumValueString(settl, r => r.Settl);
            var totalExtremaIds = BuildExtremumIdsString(total, r => r.Total);
            var settlExtremaIds = BuildExtremumIdsString(settl, r => r.Settl);

            return new GeneralReportData(
                maxTotal, minTotal, avgTotal,
                maxSettl, minSettl, avgSettl,
                noAccess, @new, destroyed,
                exceedSp, exceedCalc,
                totalExtrema, settlExtrema,
                totalExtremaIds, settlExtremaIds);
        }
    }
}
