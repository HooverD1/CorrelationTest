using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace CorrelationTest
{
    public interface IEstimateDistribution
    {
        //The distribution object should be kept private because it uses a different type (accord vs custom)
        //Interacting with the distribution object should be done via methods like GetInverse()
        //List all the common interactions in the interface

        string Name { get; set; }
        string DistributionString { get; set; }
        Dictionary<string, object> DistributionParameters { get; set; }

        double GetInverse(double percentile);
        //Dictionary<string, object> ParseStringIntoParameters { get; set; }
    }

}
