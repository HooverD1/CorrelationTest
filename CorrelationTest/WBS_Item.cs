using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace CorrelationTest
{
    public class WBS_Item : Item, IHasSubs
    {
        public Period[] Periods { get; set; }
        public UniqueID uID { get; set; }

        public WBS_Item(Excel.Range xlRow, CostSheet ContainingSheetObject) : base(xlRow, ContainingSheetObject)
        {

        }
        public List<ISub> SubEstimates { get; set; }
        public Dictionary<Estimate_Item, double> CorrelPairs { get; set; }

        public void PrintInputCorrelString() { }
        public void PrintPhasingCorrelString() { }
        public void PrintDurationCorrelString() { }
    }
}
