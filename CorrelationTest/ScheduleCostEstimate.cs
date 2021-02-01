using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace CorrelationTest
{
    public class ScheduleCostEstimate : Item, IJointEstimate
    {
        public ScheduleCostEstimate(Excel.Range xlRow, CostSheet ContainingSheetObject) : base(xlRow, ContainingSheetObject)
        {
            //This is a joint estimate
            //It needs to pull in a Cost sub-estimate
            this.costEstimate = ConstructCostSubEstimate(xlRow, ContainingSheetObject);
            this.scheduleEstimate = ConstructScheduleSubEstimate(xlRow, ContainingSheetObject);
        }
        public CostEstimate costEstimate { get; set; }
        public ScheduleEstimate scheduleEstimate { get; set; }

        public CostEstimate ConstructCostSubEstimate(Excel.Range xlRow, CostSheet ContainingSheetObject)
        {
            if (Convert.ToString(xlRow.Offset[1, ContainingSheetObject.Specs.Type_Offset].Value) != "SE")
                throw new Exception("Malformed SACE");
            int offset = Convert.ToInt32(xlRow.Offset[1, ContainingSheetObject.Specs.Level_Offset].Value) + 2;      //Number of inputs for Schedule + Schedule header + next row
            if (Convert.ToString(xlRow.Offset[offset, ContainingSheetObject.Specs.Type_Offset].Value) == "CE")
                return (CostEstimate)Item.ConstructFromRow(xlRow.Offset[offset, 0].EntireRow, ContainingSheetObject);
            else
                throw new Exception("Malformed SACE");
        }
        public ScheduleEstimate ConstructScheduleSubEstimate(Excel.Range xlRow, CostSheet ContainingSheetObject)
        {
            return (ScheduleEstimate)Item.ConstructFromRow(xlRow.Offset[1, 0].EntireRow, ContainingSheetObject);
        }
        public void LoadUID()
        {
            throw new NotImplementedException();
        }
    }
}