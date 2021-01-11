using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace CorrelationTest
{
    public class WBS_Item : Item, IHasInputSubs, IHasDurationSubs, IHasPhasingSubs
    {
        public Excel.Range xlDollarCell { get; set; }
        public Period[] Periods { get; set; }
        public int PeriodCount { get; set; } = 5;
        public List<ISub> SubEstimates { get; set; } = new List<ISub>();
        public Dictionary<Estimate_Item, double> CorrelPairs { get; set; }

        public WBS_Item(Excel.Range xlRow, CostSheet ContainingSheetObject) : base(xlRow, ContainingSheetObject)
        {

        }

        public void LoadSubEstimates()
        {
            this.SubEstimates = GetSubs();
        }

        private List<ISub> GetSubs()
        {
            throw new NotImplementedException();
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

        public void LoadPeriods()
        {
            this.Periods = GetPeriods();
        }
        private Period[] GetPeriods()
        {
            double[] dollars = LoadDollars();
            Period[] periods = new Period[PeriodCount];
            for (int i = 0; i < periods.Length; i++)
            {
                periods[i] = new Period(this.uID, $"P{i + 1}", dollars[i]);
            }
            return periods;
        }
        private double[] LoadDollars()
        {
            double[] dollars = new double[PeriodCount];
            for (int d = 0; d < dollars.Length; d++)
            {
                dollars[d] = xlDollarCell.Offset[0, d].Value ?? 0;
            }
            return dollars;
        }

        public void PrintInputCorrelString() { }
        public void PrintPhasingCorrelString() { }
        public void PrintDurationCorrelString() { }
    }
}
