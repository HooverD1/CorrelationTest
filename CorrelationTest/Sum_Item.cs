using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace CorrelationTest
{
    public class Sum_Item : Item, IHasSubs
    {
        public Period[] Periods { get; set; }
        public UniqueID uID { get; set; }

        public Sum_Item(Excel.Range xlRow, CostSheet ContainingSheetObject) : base(xlRow, ContainingSheetObject)
        {

        }
        public List<ISub> SubEstimates { get; set; }
        public Dictionary<Estimate_Item, double> CorrelPairs { get; set; }

        public List<ISub> GetSubEstimates()
        {
            throw new NotImplementedException();
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
