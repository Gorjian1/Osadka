using Osadka.Models;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Osadka.Services
{
    public record Extremum(double Value, IReadOnlyList<string> Ids);

    public record GeneralReportData(
        Extremum MaxVector, Extremum MinVector,
        Extremum MaxDx, Extremum MinDx,
        Extremum MaxDy, Extremum MinDy,
        Extremum? MaxDh, Extremum? MinDh,
        IReadOnlyList<string> NoAccessIds,
        IReadOnlyList<string> NewIds,
        IReadOnlyList<string> DestroyedIds,
        IReadOnlyList<string> ExceedVectorSpIds,
        IReadOnlyList<string> ExceedVectorCalcIds);

    public sealed class GeneralReportService
    {
        public GeneralReportData Build(
            IEnumerable<MeasurementRow> rows,
            double limitSp,
            double limitCalc)
        {
            var vec = rows.Where(r => r.Vector is { } v && !double.IsNaN(v)).ToList();
            var dxs = rows.Where(r => r.Dx is { } v && !double.IsNaN(v)).ToList();
            var dys = rows.Where(r => r.Dy is { } v && !double.IsNaN(v)).ToList();
            var dhs = rows.Where(r => r.Dh is { } v && !double.IsNaN(v)).ToList();

            Extremum GetMax(List<MeasurementRow> src, Func<MeasurementRow, double?> sel)
            {
                if (src.Count == 0) return new Extremum(double.NaN, Array.Empty<string>());
                double max = Math.Round(src.Max(r => sel(r)!.Value), 4);
                var ids = src.Where(r => Math.Abs(sel(r)!.Value - max) < 1e-6)
                             .Select(r => r.Id).ToList();
                return new Extremum(max, ids);
            }

            Extremum GetMin(List<MeasurementRow> src, Func<MeasurementRow, double?> sel)
            {
                if (src.Count == 0) return new Extremum(double.NaN, Array.Empty<string>());
                double min = Math.Round(src.Min(r => sel(r)!.Value), 4);
                var ids = src.Where(r => Math.Abs(sel(r)!.Value - min) < 1e-6)
                             .Select(r => r.Id).ToList();
                return new Extremum(min, ids);
            }

            var maxVec = GetMax(vec, r => r.Vector);
            var minVec = GetMin(vec, r => r.Vector);

            var maxDx = GetMax(dxs, r => r.Dx);
            var minDx = GetMin(dxs, r => r.Dx);

            var maxDy = GetMax(dys, r => r.Dy);
            var minDy = GetMin(dys, r => r.Dy);

            Extremum? maxDh = dhs.Count > 0 ? GetMax(dhs, r => r.Dh) : null;
            Extremum? minDh = dhs.Count > 0 ? GetMin(dhs, r => r.Dh) : null;

            // Текстовые статусы — как раньше (если используете пометки в Raw-строках)
            var noAccess = rows
                .Where(r => r.MarkRaw.Contains("нет доступ", StringComparison.OrdinalIgnoreCase))
                .Select(r => r.Id).ToList();

            var @new = rows
                .Where(r => r.SettlRaw.Contains("нов", StringComparison.OrdinalIgnoreCase))
                .Select(r => r.Id).ToList();

            var destroyed = rows
                .Where(r => r.MarkRaw.Contains("унич", StringComparison.OrdinalIgnoreCase))
                .Select(r => r.Id).ToList();

            // Превышения считаем по Вектору (СП/расчёт)
            bool Exceeded(double x, double lim) => x < -Math.Abs(lim);
            var exceedSp = vec.Where(r => Exceeded(r.Vector!.Value, limitSp)).Select(r => r.Id).ToList();
            var exceedCalc = vec.Where(r => Exceeded(r.Vector!.Value, limitCalc)).Select(r => r.Id).ToList();

            return new GeneralReportData(
                maxVec, minVec,
                maxDx, minDx,
                maxDy, minDy,
                maxDh, minDh,
                noAccess, @new, destroyed,
                exceedSp, exceedCalc);
        }
    }
}
