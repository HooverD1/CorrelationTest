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
        CostSheet ContainingSheetObject { get; set; }
        UniqueID uID { get; set; }
        void LoadUID();
    }

    public interface IHasInputSubs : IHasSubs
    {
        List<ISub> SubEstimates { get; set; }
        void LoadSubEstimates();
        Excel.Range xlCorrelCell_Inputs { get; set; }
        void PrintInputCorrelString();
    }
    public interface IHasPhasingSubs : IHasSubs
    {
        Excel.Range xlCorrelCell_Periods { get; set; }
        Excel.Range xlDollarCell { get; set; }
        Period[] Periods { get; set; }
        void LoadPeriods();
        int PeriodCount { get; set; }
        void PrintPhasingCorrelString();
    }
    public interface IHasDurationSubs : IHasSubs
    {
        void PrintDurationCorrelString();
    }
}
