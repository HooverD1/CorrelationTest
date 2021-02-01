using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace CorrelationTest
{
    public class CostEstimate : Estimate_Item, IHasInputSubs, IHasPhasingSubs, ISub
    {
        public CostEstimate(Excel.Range itemRow, CostSheet ContainingSheetObject) : base(itemRow, ContainingSheetObject)
        {
            this.CostDistributionParameters = new Dictionary<string, object>() {
                { "Type", xlDistributionCell.Offset[0,0].Value },
                { "Param1", xlDistributionCell.Offset[0,1].Value },
                { "Param2", xlDistributionCell.Offset[0,2].Value },
                { "Param3", xlDistributionCell.Offset[0,3].Value },
                { "Param4", xlDistributionCell.Offset[0,4].Value },
                { "Param5", xlDistributionCell.Offset[0,5].Value } };
            this.ValueDistribution = new Distribution(CostDistributionParameters);       //Is this useless?
        }
        public Distribution CostDistribution { get; set; }
    }
}