using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace CorrelationTest
{
    public class ScheduleCostEstimate : Estimate_Item, IHasDurationCorrelations, IHasPhasingCorrelations, IJointEstimate
    {
        public ScheduleCostEstimate(Excel.Range xlRow, CostSheet ContainingSheetObject) : base(xlRow, ContainingSheetObject)
        {
            //This is a joint estimate
            //It needs to pull in a Cost sub-estimate
            ScheduleEstimate scheduleTemplate = new ScheduleEstimate(xlRow, ContainingSheetObject);
            this.DurationCorrelationString = scheduleTemplate.DurationCorrelationString;
            this.SubEstimates = scheduleTemplate.SubEstimates;
            this.DurationDistribution = scheduleTemplate.DurationDistribution;
            this.costEstimate = ConstructCostSubEstimate(xlRow, ContainingSheetObject);
        }
        public CostEstimate costEstimate { get; set; }
        public ScheduleEstimate scheduleEstimate { get; set; }

        public CostEstimate ConstructCostSubEstimate(Excel.Range xlRow, CostSheet ContainingSheetObject)
        {
            if (Convert.ToString(xlRow.Cells[1, ContainingSheetObject.Specs.Type_Offset].Value) != "SACE")
                throw new Exception("Malformed SACE");
            if (Convert.ToString(xlRow.Cells[1,1].Offset[1, ContainingSheetObject.Specs.Type_Offset-1].Value) == "CE")
                return (CostEstimate)Item.ConstructFromRow(xlRow.Offset[1, 0].EntireRow, ContainingSheetObject);
            else
                throw new Exception("Malformed SACE");
        }
        public ScheduleEstimate ConstructScheduleSubEstimate(Excel.Range xlRow, CostSheet ContainingSheetObject)
        {
            return new ScheduleEstimate(xlRow.Offset[1, 0].EntireRow, ContainingSheetObject);
        }
       
    }
}