using System;
using System.Collections.Generic;
using System.Linq;
using Osadka.Models;

namespace Osadka.Services
{
    public record PointXY(int Cycle, double Mark);
    public record Series(string Id, IReadOnlyList<PointXY> Points);
    public sealed class DynamicsReportService
    {
        public IReadOnlyList<Series> Build(
            IReadOnlyDictionary<int, List<MeasurementRow>> cyclesDict)
        {
            var ids = cyclesDict
                .SelectMany(kv => kv.Value.Select(r => r.Id))
                .Distinct()
                .ToList();

            var list = new List<Series>();

            foreach (var id in ids)
            {
                var pts = cyclesDict
                    .OrderBy(kv => kv.Key)
                    .Select(kv =>
                    {
                        var row = kv.Value.FirstOrDefault(r => r.Id == id);

                        return row?.Total is double m
                            ? new PointXY(kv.Key, m)
                            : (PointXY?)null;
                    })
                    .Where(p => p is not null)
                    .Select(p => p!)
                    .ToList();

                list.Add(new Series(id, pts));
            }

            return list;
        }
    }
}
