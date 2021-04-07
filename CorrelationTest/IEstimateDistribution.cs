using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CorrelationTest
{
    public interface IEstimateDistribution
    {
        string Name { get; set; }
        string DistributionString { get; set; }
        Dictionary<string, object> DistributionParameters { get; set; }

        double GetInverse(double percentile);
    }
}
