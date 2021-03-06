using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace CorrelationTest
{
    public class WBS_Item : Item, IHasCostCorrelations, IHasDurationCorrelations, IHasPhasingCorrelations
    {
        public IEstimateDistribution CostDistribution { get; set; }
        public IEstimateDistribution DurationDistribution { get; set; }
        public IEstimateDistribution PhasingDistribution { get; set; }
        public Data.CorrelationString CostCorrelationString { get; set; }
        public Data.CorrelationString DurationCorrelationString { get; set; }
        public Data.CorrelationString PhasingCorrelationString { get; set; }
        public Period[] Periods { get; set; }
        public Excel.Range xlDollarCell { get; set; }
        public List<ISub> SubEstimates { get; set; } = new List<ISub>();
        public Dictionary<Estimate_Item, double> CorrelPairs { get; set; }

        public WBS_Item(Excel.Range xlRow, CostSheet ContainingSheetObject) : base(xlRow, ContainingSheetObject)
        {
            LoadPhasing(xlRow);
        }

        public void LoadSubEstimates()
        {
            this.SubEstimates = GetSubs();
        }

        private List<ISub> GetSubs()
        {
            throw new NotImplementedException();
        }

        public string[] GetFields()
        {
            IEnumerable<string> fields = from ISub sub in SubEstimates select sub.Name;
            return fields.ToArray();
        }

        public void LoadUID()
        {
            this.uID = GetUID();
        }

        protected UniqueID GetUID()
        {
            if (this.xlRow.Cells[1, ContainingSheetObject.Specs.ID_Offset].value != null)
            {
                string idString = Convert.ToString(this.xlRow.Cells[1, ContainingSheetObject.Specs.ID_Offset].value);
                return UniqueID.ConstructFromExisting(idString);
            }
            else
            {
                //Create new ID
                return UniqueID.ConstructNew("W");
            }
        }

        public void LoadPhasing(Excel.Range xlRow)
        {
            this.Periods = GetPeriods(5);
        }
        private Period[] GetPeriods(int numberOfPeriods)
        {
            double[] dollars = LoadDollars();
            Period[] periods = new Period[numberOfPeriods];
            for (int i = 0; i < periods.Length; i++)
            {
                periods[i] = new Period(this.uID, $"P{i + 1}", dollars[i]);
            }
            return periods;
        }
        private double[] LoadDollars()
        {
            double[] dollars = new double[this.Periods.Count()];
            for (int d = 0; d < dollars.Length; d++)
            {
                dollars[d] = xlDollarCell.Offset[0, d].Value ?? 0;
            }
            return dollars;
        }

        public void LoadCostCorrelString() { }
        public void PrintCostCorrelString() { }
        public void LoadPhasingCorrelString() { }
        public void PrintPhasingCorrelString() { }
        public void LoadDurationCorrelString() { }
        public void PrintDurationCorrelString() { }

        public void Expand(CorrelationType correlType)
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
