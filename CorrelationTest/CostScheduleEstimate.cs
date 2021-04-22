using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace CorrelationTest
{
    public class CostScheduleEstimate : Estimate_Item, IHasCostCorrelations, IHasPhasingCorrelations, IJointEstimate
    {
        public ScheduleEstimate scheduleEstimate { get; set; }

        public CostScheduleEstimate(Excel.Range xlRow, CostSheet ContainingSheetObject) : base(xlRow, ContainingSheetObject)
        {
            //This is a joint estimate.
            //It needs to create a sub-estimate for schedule
            //this.ValueDistributionParameters = new Dictionary<string, object>() {
            //    { "Type", xlDistributionCell.Offset[0,0].Value },
            //    { "Param1", xlDistributionCell.Offset[0,1].Value },
            //    { "Param2", xlDistributionCell.Offset[0,2].Value },
            //    { "Param3", xlDistributionCell.Offset[0,3].Value },
            //    { "Param4", xlDistributionCell.Offset[0,4].Value },
            //    { "Param5", xlDistributionCell.Offset[0,5].Value } };
            //this.CostDistribution = new SpecifiedDistribution(ValueDistributionParameters);
            this.scheduleEstimate = ConstructScheduleSubEstimate(xlRow, ContainingSheetObject);
        }
        

        public CostEstimate ConstructCostSubEstimate(Excel.Range xlRow, CostSheet ContainingSheetObject)
        {
            //CASE
            //SE
            //I
            //I
            return new CostEstimate(xlRow, ContainingSheetObject);
        }
        public ScheduleEstimate ConstructScheduleSubEstimate(Excel.Range xlRow, CostSheet ContainingSheetObject)
        {
            string CEvalue = Convert.ToString(xlRow.Cells[1, ContainingSheetObject.Specs.Type_Offset].Value);
            if (CEvalue != "CASE")
                throw new Exception("Malformed CASE");
            string SEvalue = Convert.ToString(xlRow.Cells[1, 1].Offset[1, ContainingSheetObject.Specs.Type_Offset-1].Value);
            if (SEvalue == "SE")
                return (ScheduleEstimate)Item.ConstructFromRow(xlRow.Offset[1, 0].EntireRow, ContainingSheetObject);
            else
                throw new Exception("Malformed CASE");
        }

        public override void PrintDurationCorrelString()
        {
            this.scheduleEstimate.PrintDurationCorrelString();
        }

    }
}
