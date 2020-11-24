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
        Excel.Range xlCorrelCell { get; set; }
        Dictionary<string, object> DistributionParameters { get; set; }
        string ID { get; set; }
        Excel.Range xlRow { get; set; }
        void LoadSubEstimates(Excel.Range parentRow);
    }
}
