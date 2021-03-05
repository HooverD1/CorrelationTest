using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace CorrelationTest
{
    public class CostEstimate : Estimate_Item, IHasCostCorrelations, IHasPhasingCorrelations, ISub
    {
        public CostEstimate(Excel.Range itemRow, CostSheet ContainingSheetObject) : base(itemRow, ContainingSheetObject)
        {
            this.ValueDistributionParameters = new Dictionary<string, object>() {
                { "Type", xlDistributionCell.Offset[0,0].Value },
                { "Param1", xlDistributionCell.Offset[0,1].Value },
                { "Param2", xlDistributionCell.Offset[0,2].Value },
                { "Param3", xlDistributionCell.Offset[0,3].Value },
                { "Param4", xlDistributionCell.Offset[0,4].Value },
                { "Param5", xlDistributionCell.Offset[0,5].Value } };
            this.CostDistribution = new Distribution(ValueDistributionParameters);       //Is this useless?

            //this.CorrelStringObj_Cost = Data.CorrelationString.ConstructFromParentItem_Cost(this);
            this.CorrelStringObj_Phasing = Data.CorrelationString.ConstructFromParentItem_Phasing(this);
        }
    }
}