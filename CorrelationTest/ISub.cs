using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace CorrelationTest
{
    public interface ISub
    {
        //Estimates, inputs
        DisplayCoords dispCoords { get; set; }
        Excel.Range xlTypeCell { get; set; }
        Period[] Periods { get; set; }
        UniqueID uID { get; set; }
        Distribution CostDistribution { get; set; }
        Distribution DurationDistribution { get; set; }
        string Name { get; set; }
        Dictionary<string, object> CostDistributionParameters { get; set; }
        Dictionary<string, object> DurationDistributionParameters { get; set; }
        Dictionary<Estimate_Item, double> CorrelPairs { get; set; }
    }
}
