using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace CorrelationTest
{
    namespace Data
    {
        public class CorrelationString_DM : CorrelationString
        {
            public CorrelationString_DM(string correlStringValue)
            {
                this.Value = correlStringValue;
            }

            public static CorrelationString_DM ConstructZeroString(string[] fields)
            {
                //Need to downcast csi 
                var csi = new CorrelationString(fields);
                return new CorrelationString_DM(csi.Value);
            }

            public static bool Validate()
            {
                return true;
            }

            public override void PrintToSheet(Excel.Range xlCell)
            {
                xlCell.Value = this.Value;
                xlCell.NumberFormat = "\"In Correl\";;;\"SCH_CORREL\"";
                xlCell.EntireColumn.ColumnWidth = 10;
            }
        }
    }
}
