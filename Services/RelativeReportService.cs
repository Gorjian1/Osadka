using DocumentFormat.OpenXml.Drawing;
using Osadka.Models;
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
        IReadOnlyList<RelativeRow> ExceededCalcRows,        
        Extremum MaxRelative);

    public sealed class RelativeReportService
    {
        public RelativeReport Build(
            IEnumerable<CoordRow> coords,
            IEnumerable<MeasurementRow> rows,
            double limitSp,
            double limitCalc)
        {

            var validPoints = coords
                            .Zip(rows, (coord, row) => (Coord: coord, Row: row))
                            .Where(pair => pair.Row.Total is { } t && !double.IsNaN(t))
                             .ToList();

            int n = validPoints.Count;
            var list = new List<RelativeRow>();

            for (int i = 0; i < n; i++)
            {
                for (int j = i + 1; j < n; j++)
                {

                    var(coord1, r1) = validPoints[i];
                    var(coord2, r2) = validPoints[j];

                    var(x1, y1) = (coord1.X, coord1.Y);
                    var(x2, y2) = (coord2.X, coord2.Y);

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
            Extremum maxRel;
            if (list.Count > 0)
            {
                var maxAbs = list.Max(r => Math.Abs(r.Ratio));
                var maxIds = list
                    .Where(r => Math.Abs(Math.Abs(r.Ratio) - maxAbs) < 1e-9)
                    .Select(r => $"{r.Id1}-{r.Id2}")
                    .ToList();

                maxRel = new Extremum(Math.Round(maxAbs, 6), maxIds);
            }
            else
            {
                maxRel = new Extremum(double.NaN, Array.Empty<string>());
            }
            return new RelativeReport(list, excSp, excCalc, maxRel);
        }
    }
}
