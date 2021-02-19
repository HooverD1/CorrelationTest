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
        List<ISub> SubEstimates { get; set; }
        void LoadSubEstimates();
        Excel.Range xlRow { get; set; }
        CostSheet ContainingSheetObject { get; set; }
        UniqueID uID { get; set; }
        void LoadUID();
        string[] GetFields();
    }

    public interface IHasCostSubs : IHasSubs
    {
        Data.CorrelationString CostCorrelationString { get; set; }
        Excel.Range xlCorrelCell_Cost { get; set; }
        Distribution CostDistribution { get; set; }
        void PrintCostCorrelString();
    }
    public interface IHasPhasingSubs : IHasSubs
    {
        
        Data.CorrelationString PhasingCorrelationString { get; set; }
        Excel.Range xlCorrelCell_Phasing { get; set; }
        Excel.Range xlDollarCell { get; set; }
        Period[] Periods { get; set; }      //The Periods should be the subs?
        Distribution PhasingDistribution { get; set; }
        void LoadPhasing(Excel.Range xlRow);
        void PrintPhasingCorrelString();
    }
    public interface IHasDurationSubs : IHasSubs
    {
        Data.CorrelationString DurationCorrelationString { get; set; }
        Distribution DurationDistribution { get; set; }
        Excel.Range xlCorrelCell_Duration { get; set; }
        void PrintDurationCorrelString();
    }
}
