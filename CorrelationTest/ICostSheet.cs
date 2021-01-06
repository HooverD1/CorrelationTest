using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace CorrelationTest
{
    public interface ICostSheet
    {
        Excel.Worksheet xlSheet { get; set; }
        List<IEstimate> Estimates { get; set; }
        List<IEstimate> LoadEstimates();         //loads Estimates
        bool Validate();                            //validate the fields being returned from the xlCorrelSheet against the fields in the sheetObj they're being returned to
        object[] Get_xlFields();                    //Gets the array of field names off the sheet
        void BuildCorrelations();
        Excel.Range[] PullEstimates(Excel.Range pullEstimates, Item costType);
    }
}
