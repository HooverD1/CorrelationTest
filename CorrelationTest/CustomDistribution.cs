using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CorrelationTest
{
    public class CustomDistribution : IEstimateDistribution
    {
        public string Name { get; set; }
        public string DistributionString { get; set; }
        public Dictionary<string, object> DistributionParameters { get; set; }

        public double GetInverse(double percentile)
        {
            throw new NotImplementedException();
        }
    }
}
