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
        Correlation_IT,
        Correlation_IM,
        Correlation_PT,
        Correlation_PM,
        Correlation_DT,
        Correlation_DM,
        Data,
        Model,
        FilterData,
        Input,
        Unknown
    }

    public abstract class Sheet
    {
        public Excel.Worksheet xlSheet { get; set; }

        public abstract void PrintToSheet();
        public abstract bool Validate();

        protected Excel.Range[] PullEstimates(string typeRange)       //return an array of rows
        {
            Excel.Range typeColumn = xlSheet.Range[typeRange];
            IEnumerable<Excel.Range> returnVal = from Excel.Range cell in typeColumn.Cells
                                                 where Convert.ToString(cell.Value) == "E"
                                                 select cell.EntireRow;
            return returnVal.ToArray<Excel.Range>();
        }


    }
}
