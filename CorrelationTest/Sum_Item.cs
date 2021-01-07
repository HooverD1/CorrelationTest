using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace CorrelationTest
{
    public class Sum_Item : Item, IHasInputSubs, IHasPhasingSubs, IHasDurationSubs
    {
        public Excel.Range xlDollarCell { get; set; }
        public Period[] Periods { get; set; }
        public int PeriodCount { get; set; } = 5;
        public UniqueID uID { get; set; }
        public int Level { get; set; }
        public List<ISub> SubEstimates { get; set; }
        public Dictionary<Estimate_Item, double> CorrelPairs { get; set; }

        public Sum_Item(Excel.Range xlRow, CostSheet ContainingSheetObject) : base(xlRow, ContainingSheetObject)
        {
            LoadUID();
            this.xlDollarCell = xlRow.Cells[1, ContainingSheetObject.Specs.Dollar_Offset];
            LoadPeriods();
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
                periods[i] = new Period(this.uID, i + 1, dollars[i]);
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

        public void LoadSubEstimates()
        {
            this.SubEstimates = GetSubEstimates();
        }

        private List<ISub> GetSubEstimates()
        {
            List<ISub> subEstimates = new List<ISub>();
            //Get the number of inputs
            int inputCount = Convert.ToInt32(xlRow.Cells[1, ContainingSheetObject.Specs.Level_Offset].value);    //Get the number of inputs
            for (int i = 1; i <= inputCount; i++)
            {
                subEstimates.Add(new Estimate_Item(xlRow.Offset[i, 0].EntireRow, ContainingSheetObject));
            }
            return subEstimates;
        }

        public void PrintInputCorrelString()
        {
            Data.CorrelationString inString = Data.CorrelationString.Construct(this, Data.CorrelStringType.InputsTriple);
            if (inString != null)
                inString.PrintToSheet(xlCorrelCell_Inputs);
        }
        public void PrintPhasingCorrelString()
        {
            Data.CorrelationString phString = Data.CorrelationString.Construct(this, Data.CorrelStringType.PhasingTriple);
            if (phString != null)
                phString.PrintToSheet(xlCorrelCell_Periods);
        }
        public void PrintDurationCorrelString() { }
    }
}
