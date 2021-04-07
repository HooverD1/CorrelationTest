using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace CorrelationTest
{
    public class Input_Item : Item, ISub
    {
        public IHasSubs Parent { get; set; }
        public Excel.Range xlDollarCell { get; set; }
        public DisplayCoords dispCoords { get; set; }
        public Period[] Periods { get; set; }
        public IEstimateDistribution CostDistribution { get; set; }
        public IEstimateDistribution DurationDistribution { get; set; }
        public IEstimateDistribution PhasingDistribution { get; set; }
        public Data.CorrelationString CostCorrelationString { get; set; }
        public Data.CorrelationString DurationCorrelationString { get; set; }
        public Data.CorrelationString PhasingCorrelationString { get; set; }
        public Dictionary<string, object> ValueDistributionParameters { get; set; }
        public Dictionary<string, object> PhasingDistributionParameters { get; set; }
        public Dictionary<Estimate_Item, double> CorrelPairs { get; set; }

        public Input_Item(Excel.Range xlItemRow, CostSheet ContainingSheetObject) : base(xlItemRow, ContainingSheetObject)
        {
            var specs = this.ContainingSheetObject.Specs;
            var xlDistributionCell = xlItemRow.Cells[1, specs.Distribution_Offset];

            //The Distributions should be loaded after once I can switch on the Parent property

            //this.ValueDistributionParameters = new Dictionary<string, object>() {
            //    { "Type", xlDistributionCell.Offset[0,0].Value },
            //    { "Param1", xlDistributionCell.Offset[0,1].Value },
            //    { "Param2", xlDistributionCell.Offset[0,2].Value },
            //    { "Param3", xlDistributionCell.Offset[0,3].Value },
            //    { "Param4", xlDistributionCell.Offset[0,4].Value },
            //    { "Param5", xlDistributionCell.Offset[0,5].Value } };
            //this.CostDistribution = new Distribution(ValueDistributionParameters);       //Is this useless?
            //var phasingDistributionParameters = new Dictionary<string, object>() {
            //    { "Type", "Normal" },
            //    { "Param1", 1 },
            //    { "Param2", 1 },
            //    { "Param3", 1 },
            //    { "Param4", 0 },
            //    { "Param5", 0 } };
            //this.PhasingDistribution = new Distribution(phasingDistributionParameters);    //Should this even be a Distribution object? More of a schedule.
            LoadPhasing(xlItemRow);

            this.dispCoords = DisplayCoords.ConstructDisplayCoords(SheetType.Estimate);
            this.xlCorrelCell_Cost = xlItemRow.Cells[1, dispCoords.CostCorrel_Offset];
            this.xlCorrelCell_Phasing = xlItemRow.Cells[1, dispCoords.PhasingCorrel_Offset];
            this.xlCorrelCell_Duration = xlItemRow.Cells[1, dispCoords.DurationCorrel_Offset];
        }

        public void LoadPhasing(Excel.Range xlRow)
        {
            this.Periods = GetPeriods();
        }
        private Period[] GetPeriods()       //should these be constructed as a static under Period?
        {
            Period[] periods = new Period[5];
            for (int i = 0; i < periods.Length; i++)
                periods[i] = new Period(uID, $"P{i + 1}");
            return periods;
        }

        public void LoadUID()
        {
            this.uID = GetUID();
        }
        private UniqueID GetUID()
        {
            throw new NotImplementedException();
        }
    }
}
