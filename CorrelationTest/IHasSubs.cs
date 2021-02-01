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
        Data.CorrelationString ValueCorrelationString { get; set; }
        List<ISub> SubEstimates { get; set; }
        void LoadSubEstimates();
        Excel.Range xlCorrelCell_Inputs { get; set; }
        Distribution ValueDistribution { get; set; }
        void PrintInputCorrelString();
    }
    public interface IHasPhasingSubs : IHasSubs
    {
        Data.CorrelationString PhasingCorrelationString { get; set; }
        Excel.Range xlCorrelCell_Periods { get; set; }
        Excel.Range xlDollarCell { get; set; }
        Period[] Periods { get; set; }
        Distribution PhasingDistribution { get; set; }
        void LoadPhasing(Excel.Range xlRow);
        void PrintPhasingCorrelString();
    }
    public interface IHasDurationSubs : IHasSubs
    {
        Data.CorrelationString ValueCorrelationString { get; set; }
        List<ISub> SubEstimates { get; set; }
        Distribution ValueDistribution { get; set; }
        Excel.Range xlCorrelCell_Inputs { get; set; }
        void PrintDurationCorrelString();
    }
}
