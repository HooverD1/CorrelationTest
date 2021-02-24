using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace CorrelationTest
{
    public class ScheduleEstimate : Estimate_Item, IHasDurationCorrelations, IHasPhasingCorrelations, ISub
    {
        public IHasSubs Parent { get; set; }
        public ScheduleEstimate(Excel.Range itemRow, CostSheet ContainingSheetObject) : base(itemRow, ContainingSheetObject) { }

    }
}
