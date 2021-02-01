using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace CorrelationTest
{
    public class CostScheduleEstimate : Item, IJointEstimate
    {
        public CostScheduleEstimate(Excel.Range xlRow, CostSheet ContainingSheetObject) : base(xlRow, ContainingSheetObject)
        {
            //This is a joint estimate.
            //It needs to create a sub-estimate for schedule
            this.costEstimate = ConstructCostSubEstimate(xlRow, ContainingSheetObject);
            this.scheduleEstimate = ConstructScheduleSubEstimate(xlRow, ContainingSheetObject);
        }
        public CostEstimate costEstimate { get; set; }
        public ScheduleEstimate scheduleEstimate { get; set; }

        public CostEstimate ConstructCostSubEstimate(Excel.Range xlRow, CostSheet ContainingSheetObject)
        {
            return (CostEstimate)Item.ConstructFromRow(xlRow.Offset[1, 0].EntireRow, ContainingSheetObject);
        }
        public ScheduleEstimate ConstructScheduleSubEstimate(Excel.Range xlRow, CostSheet ContainingSheetObject)
        {
            string CEvalue = Convert.ToString(xlRow.Cells[1, 1].Offset[1, ContainingSheetObject.Specs.Type_Offset-1].Value);
            if (CEvalue != "CE")
                throw new Exception("Malformed CASE");
            int offset = Convert.ToInt32(xlRow.Cells[1,1].Offset[1, ContainingSheetObject.Specs.Level_Offset-1].Value) + 2;      //Number of inputs for Cost + Cost header + next row
            string SEvalue = Convert.ToString(xlRow.Cells[1, 1].Offset[offset, ContainingSheetObject.Specs.Type_Offset - 1].Value);
            if (SEvalue == "SE")
                return (ScheduleEstimate)Item.ConstructFromRow(xlRow.Offset[offset, 0].EntireRow, ContainingSheetObject);
            else
                throw new Exception("Malformed CASE");
        }
        public void LoadUID()
        {
            throw new NotImplementedException();
        }
    }
}
