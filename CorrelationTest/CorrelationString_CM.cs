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
        public class CorrelationString_CM : CorrelationString
        {
            public CorrelationString_CM(Excel.Range xlRange) : this(GetParentID(xlRange), GetIDsFromRange(xlRange), GetFieldsFromRange(xlRange), GetCorrelArrayFromRange(xlRange)) { }

            public CorrelationString_CM(string correlString)
            {
                this.Value = ExtensionMethods.CleanStringLinebreaks(correlString);
            }

            public CorrelationString_CM(Sheets.CorrelationSheet_CM correlSheet)
            {

            }

            public CorrelationString_CM(string parent_id, object[] ids, object[] fields, object[,] correlArray)
            {
                this.Value = ExtensionMethods.CleanStringLinebreaks(CreateValue(parent_id, ids, fields, correlArray));               
            }

            private CorrelationString_CM(string parent_id, object[] sub_ids, object[] sub_fields)     //create 0 string (independence)
            {
                int fieldCount = sub_ids.Count();
                object[,] correlArray = new object[fieldCount, fieldCount];
                for(int row = 0; row < fieldCount; row++)
                {
                    for(int col = 0; col < fieldCount; col++)
                    {
                        correlArray[row, col] = 0;
                    }
                }
                this.Value = ExtensionMethods.CleanStringLinebreaks(CreateValue(parent_id, sub_ids, sub_fields, correlArray));
            }

            public CorrelationString_CM(string parent_id, object[] sub_ids, object[] sub_fields, Data.CorrelationMatrix matrix)
            {
                this.Value = ExtensionMethods.CleanStringLinebreaks(CreateValue(parent_id, sub_ids, sub_fields, matrix.GetMatrix_Values()));
            }

            private static object[] GetFieldsFromRange(Excel.Range correlRange)
            {
                var specs = new CorrelSheetSpecs(SheetType.Correlation_CM);
                Excel.Range firstCell = correlRange.Worksheet.Cells[specs.MatrixCoords.Item1, specs.MatrixCoords.Item2];
                Excel.Range lastCell = correlRange.Worksheet.Cells[specs.MatrixCoords_End.Item1, specs.MatrixCoords_End.Item2];
                Excel.Range fieldRange = correlRange.Worksheet.Range[firstCell, lastCell];
                return fieldRange.Value;
            }

            private static string GetParentID(Excel.Range correlRange)
            {
                var specs = new CorrelSheetSpecs(SheetType.Correlation_CM);
                return Convert.ToString(correlRange.Worksheet.Cells[specs.IdCoords.Item1, specs.IdCoords.Item2].value);
            }
            
            private static object[] GetIDsFromRange(Excel.Range correlRange)        //build from correlation sheet
            {
                var specs = new CorrelSheetSpecs(SheetType.Correlation_CM);
                string parentID = Convert.ToString(correlRange.Worksheet.Cells[specs.IdCoords.Item1, specs.IdCoords.Item2].value);
                string sheetID = parentID.Split('|').First();
                
                object[,] tempArray = correlRange.Resize[1, correlRange.Columns.Count].Value;
                string[] returnArray = new string[tempArray.GetLength(1)];
                for(int i = 0; i < tempArray.GetLength(1); i++)
                {
                    returnArray[i] = tempArray[1, i+1].ToString();
                }
                return returnArray;
            }

            private static object[,] GetCorrelArrayFromRange(Excel.Range correlRange)
            {
                return correlRange.Offset[1, 0].Resize[correlRange.Rows.Count - 1, correlRange.Columns.Count].Value;
            }

            private string CreateValue(IHasCostCorrelations parentEstimate)
            {
                //Convert all the sub-estimates to a correlation string
                int fields = parentEstimate.SubEstimates.Count;
                StringBuilder sb = new StringBuilder();
                //HEADER
                sb.Append($"{fields},IM");  
                for(int j = 0; j < fields; j++)
                {
                    sb.Append(",");
                    sb.Append(parentEstimate.SubEstimates[j].uID.ID);
                }
                sb.AppendLine();
                ////FIELDS
                //for(int sub = 0; sub < fields; sub++)
                //{
                //    sb.Append(parentEstimate.SubEstimates[sub].uID);
                //    if (sub < fields - 1)
                //        sb.Append(",");
                //}
                //sb.AppendLine();
                //VALUES
                for (int sub = 0; sub < fields; sub++)  //vertical
                {
                    //sb.Append(parentEstimate.SubEstimates[sub].GetID());
                    foreach(KeyValuePair<Estimate_Item, double> pair in parentEstimate.SubEstimates[sub].CorrelPairs)
                    {
                        sb.Append(pair.Value);
                        sb.Append(",");
                    }
                    sb.Remove(sb.Length - 1, 1);        //remove the trailing comma
                    if (sub < fields - 1)
                        sb.AppendLine();
                }
                return sb.ToString();
            }

            protected override string CreateValue(string parentID, object[] ids, object[] fields, object[,] correlArray)
            {
                correlArray = ExtensionMethods.ReIndexArray<object>(correlArray);
                StringBuilder sb = new StringBuilder();
                sb.Append($"{ids.Length},CM,");
                sb.Append(parentID);
                for(int i = 0; i < ids.Length; i++)
                {
                    sb.Append(",");
                    sb.Append(ids[i]);
                }
                sb.AppendLine();
                //for (int field = 0; field < fields.Length; field++)
                //{
                //    //Add fields
                //    if(field > 0)
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

            public string CreateArray(string correlString)
            {
                throw new NotImplementedException();
            }

            public static bool Validate(Excel.Range correlCell)      //validate that it is in fact a correlString
            {
                if (correlCell.NumberFormat == "\"Correl\";;;\"CORREL\"")
                    return true;
                else
                    return false;
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

            public static CorrelationString_CM ConstructZeroString(string[] fields)
            {
                //Need to downcast csi 
                var csi = new CorrelationString(fields);
                return new CorrelationString_CM(csi.Value);
            }

            //public static Data.CorrelationString_CM ConstructString(string parentID, string[] ids, object[] fields, string sheet, Dictionary<Tuple<string, string>, double> correls = null)
            //{
            //    Data.CorrelationString_CM correlationString = ConstructZeroString((from UniqueID id in ids select id.ID).ToArray());       //build zero string
            //    if (correls == null)
            //        return correlationString;       //return zero string
            //    else
            //    {
            //        Data.CorrelationMatrix matrix = Data.CorrelationMatrix.ConstructNew(correlationString);      //convert to zero matrix for modification
            //        foreach (string id1 in ids)
            //        {
            //            foreach (string id2 in ids)
            //            {
            //                if (correls.ContainsKey(new Tuple<string, string>(id1, id2)))
            //                {
            //                    matrix.SetCorrelation(id1, id2, correls[new Tuple<string, string>(id1, id2)]);
            //                }
            //                if(correls.ContainsKey(new Tuple<string, string>(id2, id1)))
            //                {
            //                    matrix.SetCorrelation(id2, id1, correls[new Tuple<string, string>(id2, id1)]);
            //                }
            //            }
            //        }
            //        //convert to a string
            //        return new Data.CorrelationString_CM(parentID, ids, fields, matrix);      //return modified zero matrix as correl string
            //    }
            //}

            //public static void ExpandCorrel(Excel.Range selection)
            //{
            //    //Verify that it's a correl string
            //    bool valid = CorrelationString_CM.Validate(selection);
            //    if (valid)
            //    {
            //        //construct the correlString
            //        Data.CorrelationString_CM correlStringObj = new Data.CorrelationString_CM(Convert.ToString(selection.Value));
            //        //construct the correlSheet
            //        Sheets.CorrelationSheet correlSheet = Sheets.CorrelationSheet.Construct(correlStringObj, selection, new Data.CorrelSheetSpecs(SheetType.Correlation_CM));
            //        //print the correlSheet
            //        correlSheet.PrintToSheet();
            //    }
            //}

            public void OverwriteIDs(UniqueID[] newIDs)
            {
                string correlString = this.Value.Replace("\r\n", "&");  //simplify delimiter
                correlString = correlString.Replace("\n", "&");  //simplify delimiter
                string[] correlLines = correlString.Split('&');         //split lines
                string[] id_strings = correlLines[0].Split(',');            //get fields (first line) and delimit
                //recombine with the newIDs
                StringBuilder sb = new StringBuilder();
                //for(int i=0; i < newIDs.Length; i++)
                //{
                //    sb.Append(newIDs[i].Name);
                //    if (i < newIDs.Length - 1)
                //        sb.Append(",");
                //}
                sb.AppendLine();
                for(int j=1;j<correlLines.Length;j++)
                {
                    sb.Append(correlLines[j]);
                }
                this.Value = sb.ToString();
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
                for (int i= 0; i < min; i++)
                {
                    xlFragments[i].Value = lines[i];
                    xlFragments[i].NumberFormat = "\"In Correl\";;;\"COST_CORREL\"";
                }
                xlFragments[0].EntireColumn.ColumnWidth = 10;
            }
        }
    }
}
