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
        public class CorrelationString_IT : CorrelationString
        {
            public Triple InputTriple { get; set; }
            public CorrelationString_IT(string correlString)
            {
                this.Value = ExtensionMethods.CleanStringLinebreaks(correlString);
                string triple = this.Value.Split('&')[2];
                this.InputTriple = new Triple(this.GetParentID().ID, triple);
            }

            public CorrelationString_IT(string[] fields, Triple it, int subs, string parent_id)        //build a triple string out of a triple
            {
                this.InputTriple = it;
                StringBuilder sb = new StringBuilder();
                sb.Append($"{subs},IT,{parent_id}");
                sb.AppendLine();
                for (int i = 0; i < fields.Length - 1; i++)
                {
                    sb.Append(fields[i]);
                    sb.Append(",");
                }
                sb.Append(fields[fields.Length-1]);
                sb.AppendLine();
                sb.Append(it.ToString());
                this.Value = ExtensionMethods.CleanStringLinebreaks(sb.ToString());
            }

            public override object[] GetFields()
            {
                string[] splitString = DelimitString();
                return splitString[1].Split(',');
            }

            public override object[,] GetMatrix()
            {
                return this.InputTriple.GetPhasingCorrelationMatrix(this.GetNumberOfPeriods()).Matrix;
            }

            public override UniqueID[] GetIDs()
            {
                string[] correlLines = DelimitString();
                string[] id_strings = correlLines[1].Split(',');            //get fields (first line) and delimit
                UniqueID[] returnIDs = id_strings.Select(x => UniqueID.ConstructFromExisting(x)).ToArray();
                if (id_strings.Distinct().Count() == id_strings.Count())
                    return returnIDs;
                else
                    throw new Exception("Duplicated IDs");
            }

            public static bool Validate()
            {
                return true;
            }

            public override UniqueID GetParentID()
            {            
                string[] lines = this.Value.Split('&');
                return UniqueID.ConstructFromExisting(lines[0]);
            }

            public override void PrintToSheet(Excel.Range xlCell)
            {
                xlCell.Value = this.Value;
                xlCell.NumberFormat = "\"In Correl\";;;\"IN_CORREL\"";
                xlCell.EntireColumn.ColumnWidth = 10;
            }

            public override void Expand(Excel.Range xlSource)
            {
                //construct the correlSheet
                Sheets.CorrelationSheet correlSheet = Sheets.CorrelationSheet.Construct(this, xlSource, new Data.CorrelSheetSpecs(SheetType.Correlation_IT));
                //print the correlSheet                         //CorrelationSheet NEEDS NEW CONSTRUCTORS BUILT FOR NON-INPUTS
                correlSheet.PrintToSheet();
            }
        }
    }
}
