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

            public CorrelationString_DM(Sheets.CorrelationSheet_DM correlSheet)
            {

            }


            public CorrelationString_DM(string parent_id, object[] sub_ids, object[] sub_fields, Data.CorrelationMatrix matrix)
            {
                this.Value = ExtensionMethods.CleanStringLinebreaks(CreateValue(parent_id, sub_ids, sub_fields, matrix.GetMatrix()));
            }

            public static bool Validate()
            {
                return true;
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

            protected override string CreateValue(string parentID, object[] ids, object[] fields, object[,] correlArray)
            {
                correlArray = ExtensionMethods.ReIndexArray<object>(correlArray);
                StringBuilder sb = new StringBuilder();
                sb.Append($"{ids.Length},DM,");
                sb.Append(parentID);
                for (int i = 0; i < ids.Length; i++)
                {
                    sb.Append(",");
                    sb.Append(ids[i]);
                }
                sb.AppendLine();
                //for (int field = 0; field < fields.Length; field++)
                //{
                //    //Add fields
                //    if (field > 0)
                //        sb.Append(",");
                //    sb.Append(fields[field]);
                //}
                //sb.AppendLine();
                for (int row = 0; row < correlArray.GetLength(0); row++)
                {
                    for (int col = row + 1; col < correlArray.GetLength(1); col++)
                    {
                        sb.Append(correlArray[row, col]);
                        if (col < correlArray.GetLength(1) - 1)
                            sb.Append(",");
                    }
                    if (row < correlArray.GetLength(0) - 2)
                        sb.AppendLine();
                }
                return sb.ToString();
            }

            public override void PrintToSheet(Excel.Range[] xlCells)
            {
                //Clean the string
                //Split the string by lines
                //Print it to the xlCells

                this.Value = ExtensionMethods.CleanStringLinebreaks(this.Value);
                List<Excel.Range> xlFragments = xlCells.ToList();
                string[] lines = this.Value.Split('&');
                int min;
                if (lines.Count() <= xlCells.Count())
                    min = lines.Count();
                else
                    min = xlCells.Count();
                for (int i = 0; i < min; i++)
                {
                    xlFragments[i].Value = lines[i];
                    xlFragments[i].NumberFormat = "\"Sch Correl\";;;\"CORREL\"";
                }
                xlFragments[0].EntireColumn.ColumnWidth = 10;
            }
        }
    }
}
