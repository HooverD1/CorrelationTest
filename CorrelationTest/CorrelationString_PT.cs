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
            public CorrelationString_PT(Triple pt, int periods, string parent_id)        //build a triple string out of a triple
            {
                this.Triple = pt;
                StringBuilder sb = new StringBuilder();
                sb.Append($"{periods},PT");
                sb.AppendLine();
                sb.Append(parent_id);
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
                //return the fields based on the parent uid and the number of periods
                return null;
                //return PeriodID.GeneratePeriodIDs(this.GetParentID(), this.GetNumberOfPeriods()).Select(x=>x.Name).ToArray<string>();
            }

            public override UniqueID[] GetIDs()
            {
                string[] correlLines = DelimitString();
                string[] id_strings = correlLines[1].Split(',');            //get fields (first line) and delimit
                UniqueID[] returnIDs = id_strings.Select(x => new UniqueID(x)).ToArray();
                if (id_strings.Distinct().Count() == id_strings.Count())
                    return returnIDs;
                else
                    throw new Exception("Duplicated IDs");
            }

            public override UniqueID GetParentID()
            {
                string[] lines = this.Value.Split('&');
                return new UniqueID(lines[1]);
            }

            public Triple GetTriple()
            {
                string[] correlLines = DelimitString();
                if (correlLines.Length != 3)
                    throw new Exception("Malformed triple string.");
                string uidString = correlLines[1];
                string tripleString = correlLines[2];
                return new Triple(uidString, tripleString);
            }
        }
    }    
}
