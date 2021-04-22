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
            //Only inputs have specified distributions -- right?
            //The cost estimate would have a custom distribution based on its inputs
            this.ValueDistributionParameters = new Dictionary<string, object>() {
                { "Type", xlDistributionCell.Offset[0,0].Value },
                { "Mean", xlDistributionCell.Offset[0,1].Value },
                { "Stdev", xlDistributionCell.Offset[0,2].Value },
                { "Param1", xlDistributionCell.Offset[0,3].Value },
                { "Param2", xlDistributionCell.Offset[0,4].Value },
                { "Param3", xlDistributionCell.Offset[0,5].Value } };
            this.CostDistribution = Distribution.ConstructForExpansion(itemRow, CorrelationType.Cost);
            this.DurationDistribution = null;
            this.PhasingDistribution = Distribution.ConstructForExpansion(itemRow, CorrelationType.Phasing);
                                                                                        //But if it can contain a custom distribution, this parameter list isn't sufficient
                                                                                          //Should all the parameters be stored as a string in a single cell?

            //this.CorrelStringObj_Cost = Data.CorrelationString.ConstructFromParentItem_Cost(this);
            this.CorrelStringObj_Phasing = Data.CorrelationString.ConstructFromParentItem_Phasing(this);
        }
    }
}