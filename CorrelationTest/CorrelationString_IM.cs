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
        public class CorrelationString_IM : CorrelationString
        {
            public CorrelationString_IM(Excel.Range xlRange) : this(GetCorrelArrayFromRange(xlRange), GetIDsFromRange(xlRange)) { }

            public CorrelationString_IM(string correlString)
            {
                this.Value = ExtensionMethods.CleanStringLinebreaks(correlString);
            }
            public CorrelationString_IM(object[,] correlArray, string[] ids)
            {
                this.Value = ExtensionMethods.CleanStringLinebreaks(CreateValue(ids, correlArray));               
            }

            private CorrelationString_IM(string[] ids, string sheet)     //create 0 string (independence)
            {
                int fieldCount = ids.Count();

                object[,] correlArray = new object[fieldCount, fieldCount];
                for(int row = 0; row < fieldCount; row++)
                {
                    for(int col = 0; col < fieldCount; col++)
                    {
                        correlArray[row, col] = 0;
                    }
                }
                this.Value = ExtensionMethods.CleanStringLinebreaks(CreateValue(ids, correlArray));
            }

            public CorrelationString_IM(Data.CorrelationMatrix matrix)
            {
                this.Value = ExtensionMethods.CleanStringLinebreaks(CreateValue(matrix.GetIDs(), matrix.GetMatrix()));
            }

            //private object[] BuildIDsFromFields(object[] fields, string sheet)
            //{
            //    object[] ids = new object[fields.Length];
            //    for (int i = 0; i < fields.Length; i++)
            //        ids[i] = $"{sheet}|{fields[i]}";
            //    return ids;
            //}
            
            private static string[] GetIDsFromRange(Excel.Range correlRange)        //build from correlation sheet
            {
                var specs = new CorrelSheetSpecs(SheetType.Correlation_IM);
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

            private string CreateValue(Estimate_Item parentEstimate)
            {
                //Convert all the sub-estimates to a correlation string
                int fields = parentEstimate.SubEstimates.Count;
                StringBuilder sb = new StringBuilder();
                sb.Append($"{fields},IM");
                sb.AppendLine();
                for(int sub = 0; sub < fields; sub++)
                {
                    sb.Append(parentEstimate.SubEstimates[sub].uID);
                    if (sub < fields - 1)
                        sb.Append(",");
                }
                sb.AppendLine();
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

            protected override string CreateValue(string[] ids, object[,] correlArray)
            {
                correlArray = ExtensionMethods.ReIndexArray<object>(correlArray);
                StringBuilder sb = new StringBuilder();
                sb.Append($"{ids.Length},IM");
                sb.AppendLine();
                for (int field = 0; field < correlArray.GetLength(1); field++)
                {
                    //Add fields
                    sb.Append(ids[field]);
                    if (field < correlArray.GetLength(1) - 1)
                        sb.Append(",");
                }
                sb.AppendLine();
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

            public override object[] GetFields()
            {
                string[] splitString = DelimitString(this.Value);
                return splitString[1].Split(',');
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

            public static CorrelationString_IM ConstructZeroString(string[] fields)
            {
                //Need to downcast csi 
                var csi = new CorrelationString(fields);
                return new CorrelationString_IM(csi.Value);
            }

            public static Data.CorrelationString_IM ConstructString(string[] ids, string sheet, Dictionary<Tuple<string, string>, double> correls = null)
            {
                Data.CorrelationString_IM correlationString = ConstructZeroString((from UniqueID id in ids select id.ID).ToArray());       //build zero string
                if (correls == null)
                    return correlationString;       //return zero string
                else
                {
                    Data.CorrelationMatrix matrix = Data.CorrelationMatrix.ConstructNew(correlationString);      //convert to zero matrix for modification
                    string[] matrixIDs = matrix.GetIDs();
                    foreach (string id1 in matrixIDs)
                    {
                        foreach (string id2 in matrixIDs)
                        {
                            if (correls.ContainsKey(new Tuple<string, string>(id1, id2)))
                            {
                                matrix.SetCorrelation(id1, id2, correls[new Tuple<string, string>(id1, id2)]);
                            }
                            if(correls.ContainsKey(new Tuple<string, string>(id2, id1)))
                            {
                                matrix.SetCorrelation(id2, id1, correls[new Tuple<string, string>(id2, id1)]);
                            }
                        }
                    }
                    //convert to a string
                    return new Data.CorrelationString_IM(matrix);      //return modified zero matrix as correl string
                }
            }

            public override void Expand(Excel.Range xlSource)
            {
                Data.CorrelationString_IM correlStringObj = new Data.CorrelationString_IM(this.Value);
                var id = this.GetIDs()[0];
                //construct the correlSheet
                Sheets.CorrelationSheet correlSheet = Sheets.CorrelationSheet.Construct(correlStringObj, xlSource, new Data.CorrelSheetSpecs(SheetType.Correlation_IM));
                //print the correlSheet                         //CorrelationSheet NEEDS NEW CONSTRUCTORS BUILT FOR NON-INPUTS
                correlSheet.PrintToSheet();
            }

            public static void ExpandCorrel(Excel.Range selection)
            {
                //Verify that it's a correl string
                bool valid = CorrelationString_IM.Validate(selection);
                if (valid)
                {
                    //construct the correlString
                    Data.CorrelationString_IM correlStringObj = new Data.CorrelationString_IM(Convert.ToString(selection.Value));
                    //construct the correlSheet
                    Sheets.CorrelationSheet correlSheet = Sheets.CorrelationSheet.Construct(correlStringObj, selection, new Data.CorrelSheetSpecs(SheetType.Correlation_IM));
                    //print the correlSheet
                    correlSheet.PrintToSheet();
                }
            }

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


            public override void PrintToSheet(Excel.Range xlCell)
            {
                xlCell.Value = this.Value;
                xlCell.NumberFormat = "\"In Correl\";;;\"IN_CORREL\"";
                xlCell.EntireColumn.ColumnWidth = 10;
            }
        }
    }
}
