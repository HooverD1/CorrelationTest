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

            public CorrelationString_IT(string[] fields, Triple it, string parent_id, string[] sub_ids)        //build a triple string out of a triple
            {
                this.InputTriple = it;
                StringBuilder sb = new StringBuilder();
                sb.Append($"{fields.Length},IT,{parent_id}");
                for (int j = 0; j < sub_ids.Length; j++)
                {
                    sb.Append(",");
                    sb.Append(sub_ids[j]);
                }
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

            public Triple GetTriple()
            {
                string[] correlLines = DelimitString(this.Value);
                if (correlLines.Length != 3)
                    throw new Exception("Malformed triple string.");
                string uidString = correlLines[0].Split(',')[2];
                string tripleString = correlLines[2];
                return new Triple(uidString, tripleString);
            }

            public override string[] GetFields()
            {
                string[] splitString = DelimitString(this.Value);
                return splitString[1].Split(',');
                //This is getting the IDs, not the fields... how to get the fields?
            }

            public override object[,] GetMatrix()
            {
                return this.InputTriple.GetCorrelationMatrix(this.GetParentID().ID, this.GetIDs(), this.GetFields(), SheetType.Correlation_IT).Matrix;
            }

            public override string[] GetIDs()
            {
                //HEADER: # INPUTS, TYPE, PARENT_ID, SUB_ID1 ... SUB_IDn
                string[] correlLines = DelimitString(this.Value);
                string[] header = correlLines[0].Split(',');            //get fields (first line) and delimit
                string parentID = header[2];
                string[] returnIDs = new string[header.Length - 3];
                for (int i = 3; i < header.Length; i++)
                    returnIDs[i - 3] = header[i];
                return returnIDs;
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
