using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace CorrelationTest
{
    public interface IEstimate
    {
        //used for Sheet types that have Correlation strings
        string Name { get; set; }
        Excel.Range xlRow { get; set; }
        Excel.Range xlNameCell { get; set; }
        Excel.Range xlCorrelCell_Inputs { get; set; }
        Excel.Range xlCorrelCell_Periods { get; set; }
        Dictionary<string, object> DistributionParameters { get; set; }
        UniqueID uID { get; set; }
        List<Estimate_Item> SubEstimates { get; set; }
        
        void LoadSubEstimates();
        void PrintName();
    }
}
