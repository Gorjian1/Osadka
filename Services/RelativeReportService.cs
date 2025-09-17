using System;
using System.Collections.Generic;
using System.Linq;
using Osadka.Models;

namespace Osadka.Services
{
    public record RelativeRow(
        string Id1,
        string Id2,
        double Distance,
        double DeltaTotal,
        double Ratio);

    // Value = |Ratio|max (модуль >= 0), Ids = все пары Id1-Id2 с таким максимумом
    public record RelativeReport(
        IReadOnlyList<RelativeRow> AllRows,
        IReadOnlyList<RelativeRow> ExceededSpRows,
        IReadOnlyList<RelativeRow> ExceededCalcRows,
        Extremum MaxRelative);

    public sealed class RelativeReportService
    {
        public RelativeReport Build(
            IEnumerable<CoordRow> coords,     // выровнены по длине с rows
            IEnumerable<MeasurementRow> rows, // rows.Count == coords.Count
            double limitSp,
            double limitCalc)
        {
            var coordList = coords.ToList();
            var dataList = rows.ToList();

            int n = Math.Min(coordList.Count, dataList.Count);
            var all = new List<RelativeRow>(n > 1 ? n * (n - 1) / 2 : 0);

            for (int i = 0; i < n; i++)
            {
                var c1 = coordList[i];
                var r1 = dataList[i];

                for (int j = i + 1; j < n; j++)
                {
                    var c2 = coordList[j];
                    var r2 = dataList[j];

                    // Координаты должны быть валидны, иначе пару пропускаем (но индексы сохраняем)
                    if (!IsFinite(c1.X) || !IsFinite(c1.Y) || !IsFinite(c2.X) || !IsFinite(c2.Y))
                        continue;

                    double dx = c2.X - c1.X;
                    double dy = c2.Y - c1.Y;
                    double dist = Math.Sqrt(dx * dx + dy * dy);
                    if (!(dist > 1e-9)) // защита от 0 и NaN/Infinity
                        continue;

                    // Пара считается только если у обеих точек Total числовой
                    double t1 = r1.Total ?? double.NaN;
                    double t2 = r2.Total ?? double.NaN;
                    if (double.IsNaN(t1) || double.IsNaN(t2))
                        continue;

                    double dT = t2 - t1;
                    double ratio = dT / dist;

                    all.Add(new RelativeRow(
                        r1.Id, r2.Id,
                        Math.Round(dist, 4),
                        Math.Round(dT, 4),
                        Math.Round(ratio, 6)));
                }
            }

            var excSp = all.Where(r => Math.Abs(r.Ratio) > limitSp).ToList();
            var excCalc = all.Where(r => Math.Abs(r.Ratio) > limitCalc).ToList();

            // Максимум по модулю |Ratio|
            Extremum maxRel;
            if (all.Count > 0)
            {
                var maxAbs = all.Max(r => Math.Abs(r.Ratio));
                var maxIds = all
                    .Where(r => Math.Abs(Math.Abs(r.Ratio) - maxAbs) < 1e-9)
                    .Select(r => $"{r.Id1}-{r.Id2}")
                    .ToList();

                maxRel = new Extremum(Math.Round(maxAbs, 6), maxIds);
            }
            else
            {
                maxRel = new Extremum(double.NaN, Array.Empty<string>());
            }

            return new RelativeReport(all, excSp, excCalc, maxRel);
        }

        private static bool IsFinite(double v) => !(double.IsNaN(v) || double.IsInfinity(v));
    }
}
