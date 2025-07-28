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

    public record RelativeReport(
        IReadOnlyList<RelativeRow> AllRows,
        IReadOnlyList<RelativeRow> ExceededSpRows,
        IReadOnlyList<RelativeRow> ExceededCalcRows);

    public sealed class RelativeReportService
    {
        public RelativeReport Build(
            IEnumerable<CoordRow> coords,
            IEnumerable<MeasurementRow> rows,
            double limitSp,
            double limitCalc)
        {
            var coordList = coords.ToList();
            var dataList = rows
                .Where(r => r.Total is { } t && !double.IsNaN(t))
                .ToList();

            int n = Math.Min(coordList.Count, dataList.Count);
            var list = new List<RelativeRow>();

            for (int i = 0; i < n; i++)
            {
                for (int j = i + 1; j < n; j++)
                {
                    var r1 = dataList[i];
                    var r2 = dataList[j];

                    var (x1, y1) = (coordList[i].X, coordList[i].Y);
                    var (x2, y2) = (coordList[j].X, coordList[j].Y);

                    double dx = x2 - x1;
                    double dy = y2 - y1;
                    double dist = Math.Sqrt(dx * dx + dy * dy);

                    double dT = r2.Total!.Value - r1.Total!.Value;
                    double ratio = dist > 0 ? dT / dist : double.NaN;

                    list.Add(new RelativeRow(
                        r1.Id, r2.Id,
                        Math.Round(dist, 4),
                        Math.Round(dT, 4),
                        Math.Round(ratio, 6)));
                }
            }

            var excSp = list.Where(r => Math.Abs(r.Ratio) > limitSp).ToList();
            var excCalc = list.Where(r => Math.Abs(r.Ratio) > limitCalc).ToList();

            return new RelativeReport(list, excSp, excCalc);
        }
    }
}
