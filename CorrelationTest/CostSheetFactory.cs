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
        
    }

    public static class CostSheetFactory       
    {                                   
        public static ICostSheet Construct(Dictionary<string, object> data)
        {
            ICostSheet sheetObj;
            SheetType sheetType = (SheetType)data["SheetType"];
            switch (sheetType)
            {
                case SheetType.WBS:
                    sheetObj = new Sheets.WBSSheet((Excel.Worksheet)data["xlSheet"]);
                    break;
                case SheetType.Estimate:
                    sheetObj = new Sheets.EstimateSheet((Excel.Worksheet)data["xlSheet"]);
                    break;
                default:
                    sheetObj = null;
                    break;
            }
            return sheetObj;
        }
    }
}
