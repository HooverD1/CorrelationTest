using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace CorrelationTest
{
    public class Input_Item : Item, ISub
    {
        public DisplayCoords dispCoords { get; set; }
        public int PeriodCount { get; set; }
        public Period[] Periods { get; set; }
        public UniqueID uID { get; set; }
        public Distribution ItemDistribution { get; set; }
        public Dictionary<string, object> DistributionParameters { get; set; }
        public Dictionary<Estimate_Item, double> CorrelPairs { get; set; }

        public Input_Item(Excel.Range xlRow, CostSheet ContainingSheetObject) : base(xlRow, ContainingSheetObject)
        {

        }
        
    }
}
