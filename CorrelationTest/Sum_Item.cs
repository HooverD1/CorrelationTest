using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace CorrelationTest
{
    public class Sum_Item : Estimate_Item, IHasCostCorrelations, IHasPhasingCorrelations, IHasDurationCorrelations
    {
        
        public Sum_Item(Excel.Range xlItemRow, CostSheet ContainingSheetObject) : base(xlItemRow, ContainingSheetObject)
        {
            LoadUID();
            this.xlDollarCell = xlItemRow.Cells[1, ContainingSheetObject.Specs.Dollar_Offset];
            LoadPhasing(xlItemRow);


            this.ValueDistributionParameters = new Dictionary<string, object>() {
                { "Type", xlDistributionCell.Offset[0,0].Value },
                { "Mean", xlDistributionCell.Offset[0,1].Value },
                { "Stdev", xlDistributionCell.Offset[0,2].Value },
                { "Param1", xlDistributionCell.Offset[0,3].Value },
                { "Param2", xlDistributionCell.Offset[0,4].Value },
                { "Param3", xlDistributionCell.Offset[0,5].Value } };

            this.dispCoords = DisplayCoords.ConstructDisplayCoords(SheetType.WBS);
            this.xlCorrelCell_Cost = xlItemRow.Cells[1, dispCoords.CostCorrel_Offset];
            this.xlCorrelCell_Phasing = xlItemRow.Cells[1, dispCoords.PhasingCorrel_Offset];
            this.xlCorrelCell_Duration = xlItemRow.Cells[1, dispCoords.DurationCorrel_Offset];
        }

        protected override UniqueID GetUID()
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

        //public override string[] GetFields()
        //{
        //    IEnumerable<string> fields = from ISub sub in SubEstimates select sub.Name;
        //    return fields.ToArray();
        //}

        //public override void LoadPhasing(Excel.Range xlRow)
        //{

        //    this.PhasingDistribution = Distribution.ConstructForExpansion(xlRow, CorrelationType.Phasing);       //distribution cost and schedule distributions need differentiated
        //    this.Periods = GetPeriods();
        //}
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

        //public void LoadSubEstimates()
        //{
        //    this.SubEstimates = GetSubs();
        //}
        
        protected override List<ISub> GetSubs()
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

        public override void LoadCostCorrelString()
        {
            if (!this.SubEstimates.Any())
                return;
            var firstChild = this.SubEstimates.First();
            if (firstChild.xlCorrelCell_Cost.Value == null)
            {
                //This should check the correl type to split pairwise from matrix
                this.CostCorrelationString = Data.CorrelationString.ConstructDefaultFromCostSheet(this, Data.CorrelStringType.CostPair);
            }
            else
            {
                //Something in the cell that can either be resolved into a correl string or not
                try
                {
                    string costString = Data.CorrelationString.ConstructStringFromRange(from ISub sub in this.SubEstimates select sub.xlCorrelCell_Cost);
                    this.CostCorrelationString = Data.CorrelationString.ConstructFromStringValue(costString);
                }
                catch
                {
                    if (MyGlobals.DebugMode)
                        throw new Exception("Malformed correl string");
                }
            }
            //this.CostCorrelationString = Data.CorrelationString.ConstructDefaultFromCostSheet(this, Data.CorrelStringType.CostPair);
        }

        public override void PrintCostCorrelString()
        {
            IEnumerable<Excel.Range> xlFragments = from ISub sub in this.SubEstimates select sub.xlCorrelCell_Cost;
            if (this.CostCorrelationString != null)
                this.CostCorrelationString.PrintToSheet(xlFragments.ToArray());
        }

        public override void LoadPhasingCorrelString()
        {
            this.PhasingCorrelationString = Data.CorrelationString.ConstructDefaultFromCostSheet(this, Data.CorrelStringType.PhasingPair);
        }

        public override void PrintPhasingCorrelString()
        {
            if (this.PhasingCorrelationString != null)
                this.PhasingCorrelationString.PrintToSheet(xlCorrelCell_Phasing);
        }

        public override void LoadDurationCorrelString()
        {
            this.DurationCorrelationString = Data.CorrelationString.ConstructDefaultFromCostSheet(this, Data.CorrelStringType.DurationPair);
        }

        public override void PrintDurationCorrelString()
        {
            IEnumerable<Excel.Range> xlFragments = from ISub sub in this.SubEstimates
                                                   select sub.xlCorrelCell_Duration;
            if (this.DurationCorrelationString != null)
                this.DurationCorrelationString.PrintToSheet(xlFragments.ToArray());
        }

        public override void Expand(CorrelationType correlType)
        {
            switch (correlType)
            {
                case CorrelationType.Cost:
                    Expand_Cost();
                    break;
                case CorrelationType.Phasing:
                    Expand_Phasing();
                    break;
                case CorrelationType.Duration:
                    Expand_Duration();
                    break;
            }
        }

        private void Expand_Cost()
        {
            SheetType typeOfCost = this.CostCorrelationString.GetCorrelType();
            Sheets.CorrelationSheet correlSheet = Sheets.CorrelationSheet.ConstructFromParentItem(this, typeOfCost);
            correlSheet.PrintToSheet();
        }

        private void Expand_Phasing()
        {
            SheetType typeOfCost = this.PhasingCorrelationString.GetCorrelType();
            Sheets.CorrelationSheet correlSheet = Sheets.CorrelationSheet.ConstructFromParentItem(this, typeOfCost);
            correlSheet.PrintToSheet();
        }

        private void Expand_Duration()
        {
            SheetType typeOfCost = this.DurationCorrelationString.GetCorrelType();
            Sheets.CorrelationSheet correlSheet = Sheets.CorrelationSheet.ConstructFromParentItem(this, typeOfCost);
            correlSheet.PrintToSheet();
        }
    }
}
