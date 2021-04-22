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
        //I'm not sure Schedule Estimates are a thing.. If you need to estimate the schedule it should be a CASE or SACE, no?

        public ScheduleEstimate(Excel.Range itemRow, CostSheet ContainingSheetObject) : base(itemRow, ContainingSheetObject)
        {
            this.CostDistribution = null;
            this.DurationDistribution = Distribution.ConstructForExpansion(itemRow, CorrelationType.Duration);
            this.PhasingDistribution = null;

            this.ValueDistributionParameters = new Dictionary<string, object>() {
                { "Type", xlDistributionCell.Offset[0,0].Value },
                { "Mean", xlDistributionCell.Offset[0,1].Value },
                { "Stdev", xlDistributionCell.Offset[0,2].Value },
                { "Param1", xlDistributionCell.Offset[0,3].Value },
                { "Param2", xlDistributionCell.Offset[0,4].Value },
                { "Param3", xlDistributionCell.Offset[0,5].Value } };
        }

    }
}
