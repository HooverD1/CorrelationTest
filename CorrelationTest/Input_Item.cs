using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace CorrelationTest
{
    public class Input_Item : Item, ISub, IHasPhasingSubs
    {
        public Excel.Range xlDollarCell { get; set; }
        public DisplayCoords dispCoords { get; set; }
        public Period[] Periods { get; set; }
        public Distribution CostDistribution { get; set; }
        public Distribution DurationDistribution { get; set; }
        public Distribution PhasingDistribution { get; set; }
        public Data.CorrelationString CostCorrelationString { get; set; }
        public Data.CorrelationString DurationCorrelationString { get; set; }
        public Data.CorrelationString PhasingCorrelationString { get; set; }
        public Dictionary<string, object> CostDistributionParameters { get; set; }
        public Dictionary<string, object> DurationDistributionParameters { get; set; }
        public Dictionary<string, object> PhasingDistributionParameters { get; set; }
        public Dictionary<Estimate_Item, double> CorrelPairs { get; set; }

        public Input_Item(Excel.Range xlRow, CostSheet ContainingSheetObject) : base(xlRow, ContainingSheetObject)
        {
            var specs = this.ContainingSheetObject.Specs;
            var xlDistributionCell = xlRow.Cells[1, specs.Distribution_Offset];
            var costDistributionParameters = new Dictionary<string, object>() {
                { "Type", xlDistributionCell.Offset[0,0].Value },
                { "Param1", xlDistributionCell.Offset[0,1].Value },
                { "Param2", xlDistributionCell.Offset[0,2].Value },
                { "Param3", xlDistributionCell.Offset[0,3].Value },
                { "Param4", xlDistributionCell.Offset[0,4].Value },
                { "Param5", xlDistributionCell.Offset[0,5].Value } };
            this.CostDistribution = new Distribution(costDistributionParameters);       //Is this useless?
            var phasingDistributionParameters = new Dictionary<string, object>() {
                { "Type", "Normal" },
                { "Param1", 1 },
                { "Param2", 1 },
                { "Param3", 1 },
                { "Param4", 0 },
                { "Param5", 0 } };
            this.PhasingDistribution = new Distribution(phasingDistributionParameters);    //Should this even be a Distribution object? More of a schedule.
            LoadPhasing(xlRow);
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

        public void PrintPhasingCorrelString()
        {
            Data.CorrelationString phString = Data.CorrelationString.ConstructNew(this, Data.CorrelStringType.PhasingTriple);
            if (phString != null)
                phString.PrintToSheet(xlCorrelCell_Periods);
        }
    }
}
