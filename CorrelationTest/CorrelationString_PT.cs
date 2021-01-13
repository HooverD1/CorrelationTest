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
        public class CorrelationString_PT : CorrelationString
        {
            public Triple Triple { get; set; }
            public CorrelationString_PT(Triple pt, string[] start_dates, string parent_id)        //build a triple string out of a triple
            {
                this.Triple = pt;
                StringBuilder sb = new StringBuilder();
                sb.Append($"{start_dates.Length},PT,{parent_id}");  //Header
                sb.AppendLine();
                foreach (string start_date in start_dates)
                {
                    sb.Append(start_date);    //Period start dates as fields
                    sb.Append(",");
                }
                sb.Remove(sb.Length - 1, 1);    //remove the final comma on the fields
                sb.AppendLine();
                sb.Append(pt.ToString());
                this.Value = ExtensionMethods.CleanStringLinebreaks(sb.ToString());
            }

            public CorrelationString_PT(string correlString)
            {
                this.Value = ExtensionMethods.CleanStringLinebreaks(correlString);
                string triple = this.Value.Split('&')[2];
                this.Triple = new Triple(this.GetParentID().ID, triple);
            }

            public override void Expand(Excel.Range xlSource)
            {
                //construct the correlSheet
                Sheets.CorrelationSheet correlSheet = Sheets.CorrelationSheet.Construct(this, xlSource, new Data.CorrelSheetSpecs(SheetType.Correlation_PT));
                //print the correlSheet                         //CorrelationSheet NEEDS NEW CONSTRUCTORS BUILT FOR NON-INPUTS
                correlSheet.PrintToSheet();
            }

            public override object[,] GetMatrix()
            {
                return this.Triple.GetPhasingCorrelationMatrix(this.GetNumberOfPeriods()).Matrix;
            }

            public override void PrintToSheet(Excel.Range xlCell)
            {
                xlCell.Value = this.Value;
                xlCell.NumberFormat = "\"Ph Correl\";;;\"PH_CORREL\"";
                xlCell.EntireColumn.ColumnWidth = 10;
            }

            public override object[] GetFields()
            {
                //HEADER [Array size,Correl type,ParentID]
                //FIELDS [Field1,Field2, ... ,Field n]      //Store start dates as fields for PT
                //VALUES [0,0,0]
                string[] lines = DelimitString(this.Value);
                string[] header = lines[0].Split(',');                
                string[] fields = lines[1].Split(',');
                if (!int.TryParse(Convert.ToString(header[0]), out int size)) { throw new Exception("Malformed Correlation String"); }
                if (size != fields.Length) { throw new Exception("Malformed Correlation String"); }
                return fields.ToArray<object>();
            }

            public override UniqueID GetParentID()
            {
                string[] lines = CorrelationString.DelimitString(this.Value);
                string[] header = lines[0].Split(',');
                return UniqueID.ConstructFromExisting(header[2]);
            }

            public Triple GetTriple()
            {
                string[] correlLines = DelimitString(this.Value);
                if (correlLines.Length != 3)
                    throw new Exception("Malformed triple string.");
                string uidString = correlLines[0].Split(',')[2];
                string tripleString = correlLines[2];
                return new Triple(uidString, tripleString);
            }

            public override string[] GetIDs()
            {
                var period_ids = PeriodID.GeneratePeriodIDs(this.GetParentID(), this.GetNumberOfPeriods());
                return period_ids.Select(x => x.ID).ToArray();
            }

            public static bool Validate()
            {
                return true;
            }
        }
    }    
}
