using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace CorrelationTest
{

    public enum SheetType
    {
        WBS,
        Estimate,
        Correlation,
        Data,
        Model,
        FilterData,
        Input,
        Unknown
    }

    namespace Sheets
    {
        public abstract class Sheet
        {
            public Excel.Worksheet xlSheet { get; set; }

            public abstract void PrintToSheet();
            public abstract bool Validate();

            public static SheetType GetSheetType(Excel.Worksheet xlSheet)
            {
                string sheetIdent = xlSheet.Cells[1, 1].Value;
                switch (sheetIdent)
                {
                    case "$Correlation":
                        return SheetType.Correlation;
                    case "$WBS":
                        return SheetType.WBS;
                    case "$Estimate":
                        return SheetType.Estimate;
                    default:
                        return SheetType.Unknown;
                }
            }
        }
    }
}
