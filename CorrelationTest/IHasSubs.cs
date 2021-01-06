using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace CorrelationTest
{
    public interface IHasSubs
    {
        Excel.Range xlRow { get; set; }
        Period[] Periods { get; set; }
        CostSheet ContainingSheetObject { get; set; }
        UniqueID uID { get; set; }
        Excel.Range xlCorrelCell_Inputs { get; set; }
        Excel.Range xlCorrelCell_Periods { get; set; }
        List<ISub> SubEstimates { get; set; }
        void PrintInputCorrelString();
        void PrintPhasingCorrelString();
        void PrintDurationCorrelString();

    }
}
