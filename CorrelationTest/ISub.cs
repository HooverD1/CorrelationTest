using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CorrelationTest
{
    public interface ISub
    {
        //Estimates, inputs
        DisplayCoords dispCoords { get; set; }
        int PeriodCount { get; set; }
        Period[] Periods { get; set; }
        UniqueID uID { get; set; }
        Distribution ItemDistribution { get; set; }
        Dictionary<string, object> DistributionParameters { get; set; }
        Dictionary<Estimate_Item, double> CorrelPairs { get; set; }
    }
}
