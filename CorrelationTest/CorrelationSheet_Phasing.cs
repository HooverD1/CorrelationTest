using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace CorrelationTest
{
    namespace Sheets
    {
        public class CorrelationSheet_Phasing : CorrelationSheet
        {
            public CorrelationSheet_Phasing(Data.CorrelationString_Periods correlString, Excel.Range launchedFrom, Data.CorrelSheetSpecs specs) : base(correlString, launchedFrom, specs)
            {

            }


        }
    }
}
