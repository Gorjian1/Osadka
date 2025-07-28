using Osadka.Models;
using Osadka.ViewModels;
using System.Collections.Generic;

namespace Osadka.Models
{
    public class ProjectData
    {
        public int Cycle { get; set; }
        public double? MaxNomen { get; set; }
        public double? MaxCalculated { get; set; }
        public double? RelNomen { get; set; }
        public double? RelCalculated { get; set; }

        public List<MeasurementRow> DataRows { get; set; } = new();
        public List<CoordRow> CoordRows { get; set; } = new();

   public Dictionary<int, Dictionary<int, List<MeasurementRow>>> Objects { get; set; }       = new ();
}

}
