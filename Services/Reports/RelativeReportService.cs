using System;
using System.Collections.Generic;
using System.Linq;
using Osadka.Models;

namespace Osadka.Services.Reports
{
    public record RelativeRow(
        string Id1,
        string Id2,
        double Distance,    // мм
        double DeltaTotal,  // мм
        double Ratio);      // ΔS / Dist

   // public record Extremum(double Value, IReadOnlyList<string> Ids);

    public record RelativeReport(
        IReadOnlyList<RelativeRow> AllRows,
        IReadOnlyList<RelativeRow> ExceededSpRows,
        IReadOnlyList<RelativeRow> ExceededCalcRows,
        Extremum MaxRelative);

    public sealed class RelativeReportService
    {
        public RelativeReport Build(
            IEnumerable<CoordRow> coordsAligned,   // ВЫРАВНЕНЫ ПО ПОЗИЦИИ под DataRows
            IEnumerable<MeasurementRow> rows,      // та же длина, что coordsAligned
            double limitSp,
            double limitCalc)
        {
            // НИКАКОЙ фильтрации по Total — пары строим строго по индексам!
            var points = coordsAligned.Zip(rows, (c, r) => (Coord: c, Row: r)).ToList();

            int n = points.Count;
            var all = new List<RelativeRow>(n > 1 ? n * (n - 1) / 2 : 0);

            for (int i = 0; i < n; i++)
            {
                for (int j = i + 1; j < n; j++)
                {
                    var (c1, r1) = points[i];
                    var (c2, r2) = points[j];

                    // Дистанция — только если у обеих точек валидные координаты
                    double dist;
                    if (IsFinite(c1.X) && IsFinite(c1.Y) && IsFinite(c2.X) && IsFinite(c2.Y))
                    {
                        double dx = c2.X - c1.X;
                        double dy = c2.Y - c1.Y;
                        dist = Math.Sqrt(dx * dx + dy * dy);
                    }
                    else
                    {
                        dist = double.NaN;
                    }

                    // ΔS — только если оба Total заданы
                    double dS;
                    if (r1.Total is double t1 && IsFinite(t1) &&
                        r2.Total is double t2 && IsFinite(t2))
                    {
                        dS = t2 - t1;
                    }
                    else
                    {
                        dS = double.NaN;
                    }

                    // Относительная — только при валидных dist и dS
                    double ratio = (IsFinite(dist) && dist > 0 && IsFinite(dS))
                        ? dS / dist
                        : double.NaN;

                    all.Add(new RelativeRow(
                        r1.Id, r2.Id,
                        RoundOrNaN(dist, 4),
                        RoundOrNaN(dS, 4),
                        RoundOrNaN(ratio, 6)));
                }
            }

            // Превышения считаем только по валидным ratio
            var valid = all.Where(r => IsFinite(r.Ratio)).ToList();
            var excSp = valid.Where(r => Math.Abs(r.Ratio) > limitSp).ToList();
            var excCalc = valid.Where(r => Math.Abs(r.Ratio) > limitCalc).ToList();

            // Максимум по модулю
            Extremum maxRel;
            if (valid.Count > 0)
            {
                var maxAbs = valid.Max(r => Math.Abs(r.Ratio));
                var ids = valid
                    .Where(r => Math.Abs(Math.Abs(r.Ratio) - maxAbs) < 1e-9)
                    .Select(r => $"{r.Id1}-{r.Id2}")
                    .ToList();

                maxRel = new Extremum(Math.Round(maxAbs, 6), ids);
            }
            else
            {
                maxRel = new Extremum(double.NaN, Array.Empty<string>());
            }

            return new RelativeReport(all, excSp, excCalc, maxRel);
        }

        private static bool IsFinite(double v) => !double.IsNaN(v) && !double.IsInfinity(v);

        private static double RoundOrNaN(double v, int digits)
            => IsFinite(v) ? Math.Round(v, digits) : double.NaN;
    }
}
