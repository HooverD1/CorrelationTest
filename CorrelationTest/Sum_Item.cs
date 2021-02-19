using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace CorrelationTest
{
    public class Sum_Item : Item, IHasCostSubs, IHasPhasingSubs, IHasDurationSubs
    {
        public DisplayCoords dispCoords { get; set; }
        public Excel.Range xlDollarCell { get; set; }
        public Period[] Periods { get; set; }
        public Distribution CostDistribution { get; set; }
        public Distribution PhasingDistribution { get; set; }
        public Distribution DurationDistribution { get; set; }
        public Data.CorrelationString CostCorrelationString { get; set; }
        public Data.CorrelationString PhasingCorrelationString { get; set; }
        public Data.CorrelationString DurationCorrelationString { get; set; }
        public List<ISub> SubEstimates { get; set; } = new List<ISub>();
        public Dictionary<Estimate_Item, double> CorrelPairs { get; set; }

        public Sum_Item(Excel.Range xlItemRow, CostSheet ContainingSheetObject) : base(xlItemRow, ContainingSheetObject)
        {
            LoadUID();
            this.xlDollarCell = xlItemRow.Cells[1, ContainingSheetObject.Specs.Dollar_Offset];
            LoadPhasing(xlItemRow);

            this.dispCoords = DisplayCoords.ConstructDisplayCoords(SheetType.WBS);
            this.xlCorrelCell_Cost = xlItemRow.Cells[1, dispCoords.CostCorrel_Offset];
            this.xlCorrelCell_Phasing = xlItemRow.Cells[1, dispCoords.PhasingCorrel_Offset];
            this.xlCorrelCell_Duration = xlItemRow.Cells[1, dispCoords.DurationCorrel_Offset];
        }

        public void LoadUID()
        {
            this.uID = GetUID();
        }

        private UniqueID GetUID()
        {
            if(this.xlRow.Cells[1, ContainingSheetObject.Specs.ID_Offset].value != null)
            {
                string idString = Convert.ToString(this.xlRow.Cells[1, ContainingSheetObject.Specs.ID_Offset].value);
                return UniqueID.ConstructFromExisting(idString);                
            }
            else
            {
                //Create new ID
                return UniqueID.ConstructNew("S");
            }
        }

        public string[] GetFields()
        {
            IEnumerable<string> fields = from ISub sub in SubEstimates select sub.Name;
            return fields.ToArray();
        }

        public void LoadPhasing(Excel.Range xlRow)
        {
            var phasingDistributionParameters = new Dictionary<string, object>() {
                { "Type", "Normal" },
                { "Param1", 1 },
                { "Param2", 1 },
                { "Param3", 1 },
                { "Param4", 0 },
                { "Param5", 0 } };
            this.PhasingDistribution = new Distribution(phasingDistributionParameters);
            this.Periods = GetPeriods();
        }
        private Period[] GetPeriods()
        {
            double[] dollars = LoadDollars();
            Period[] periods = new Period[5];
            for (int i = 0; i < periods.Length; i++)
            {
                periods[i] = new Period(this.uID, $"P{i + 1}", dollars[i]);
            }
            return periods;
        }
        private double[] LoadDollars()
        {
            double[] dollars = new double[5];
            for (int d = 0; d < dollars.Length; d++)
            {
                dollars[d] = xlDollarCell.Offset[0, d].Value ?? 0;
            }
            return dollars;
        }

        public void LoadSubEstimates()
        {
            this.SubEstimates = GetSubs();
        }
        
        private List<ISub> GetSubs()
        {
            List<ISub> subEstimates = new List<ISub>();
            //Get the number of inputs
            int inputCount = Convert.ToInt32(xlRow.Cells[1, ContainingSheetObject.Specs.Level_Offset].value);    //Get the number of inputs
            for (int i = 1; i <= inputCount; i++)
            {
                subEstimates.Add((ISub)Item.ConstructFromRow(xlRow.Offset[i, 0].EntireRow, ContainingSheetObject));
            }
            return subEstimates;
        }

        public void PrintCostCorrelString()
        {
            Data.CorrelationString inString = Data.CorrelationString.ConstructDefaultFromCostSheet(this, Data.CorrelStringType.CostTriple);
            IEnumerable<Excel.Range> xlFragments = from ISub sub in this.SubEstimates
                                                   select sub.xlCorrelCell_Cost;
            if (inString != null)
                inString.PrintToSheet(xlFragments.ToArray());
        }
        public void PrintPhasingCorrelString()
        {
            Data.CorrelationString phString = Data.CorrelationString.ConstructDefaultFromCostSheet(this, Data.CorrelStringType.PhasingTriple);
            if (phString != null)
                phString.PrintToSheet(xlCorrelCell_Phasing.Resize[1,this.Periods.Count()]);
        }
        public void PrintDurationCorrelString()
        {
            Data.CorrelationString phString = Data.CorrelationString.ConstructDefaultFromCostSheet(this, Data.CorrelStringType.DurationTriple);
            IEnumerable<Excel.Range> xlFragments = from ISub sub in this.SubEstimates
                                                   select sub.xlCorrelCell_Duration;
            if (phString != null)
                phString.PrintToSheet(xlFragments.ToArray());
        }
    }
}
