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
        List<IHasSubs> Parents { get; set; }
        DisplayCoords dispCoords { get; set; }
        Excel.Range xlTypeCell { get; set; }
        Excel.Range xlCorrelCell_Cost { get; set; }
        Excel.Range xlCorrelCell_Duration { get; set; }
        Excel.Range xlCorrelCell_Phasing { get; set; }
        Period[] Periods { get; set; }
        UniqueID uID { get; set; }
        Distribution CostDistribution { get; set; }
        Distribution DurationDistribution { get; set; }
        string Name { get; set; }
        Dictionary<string, object> ValueDistributionParameters { get; set; }
        Dictionary<Estimate_Item, double> CorrelPairs { get; set; }
    }
}
