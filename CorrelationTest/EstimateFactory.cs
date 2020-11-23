using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace CorrelationTest
{
    public static class EstimateFactory
    {
        public static IEstimate Construct(Excel.Range selection)
        {
            SheetType sheetType = Sheets.Sheet.GetSheetType(selection.Worksheet);
            IEstimate returnEstimate;
            switch (sheetType)
            {
                case SheetType.WBS:
                    returnEstimate = new Estimate(selection.EntireRow);
                    break;
                case SheetType.Unknown:
                    throw new NotImplementedException();
                default:
                    throw new NotImplementedException();
            };
            return returnEstimate;
        }
    }
}
