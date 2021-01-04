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
            public CorrelationString_IM(object[,] correlArray, UniqueID[] ids)
            {
                this.Value = ExtensionMethods.CleanStringLinebreaks(CreateValue(ids, correlArray));               
            }

            private CorrelationString_IM(UniqueID[] ids, string sheet)     //create 0 string (independence)
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
            
            private static UniqueID[] GetIDsFromRange(Excel.Range correlRange)        //build from correlation sheet
            {
                var specs = new CorrelSheetSpecs(SheetType.Correlation_IM);
                string parentID = Convert.ToString(correlRange.Worksheet.Cells[specs.IdCoords.Item1, specs.IdCoords.Item2].value);
                string sheetID = parentID.Split('|').First();
                
                object[,] tempArray = correlRange.Resize[1, correlRange.Columns.Count].Value;
                UniqueID[] returnArray = new UniqueID[tempArray.GetLength(1)];
                for(int i = 0; i < tempArray.GetLength(1); i++)
                {
                    returnArray[i] = new UniqueID(sheetID, tempArray[1, i+1].ToString());
                }
                return returnArray;
            }
            private static object[,] GetCorrelArrayFromRange(Excel.Range correlRange)
            {
                return correlRange.Offset[1, 0].Resize[correlRange.Rows.Count - 1, correlRange.Columns.Count].Value;
            }

            private string CreateValue(Estimate parentEstimate)
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
                    foreach(KeyValuePair<Estimate, double> pair in parentEstimate.SubEstimates[sub].CorrelPairs)
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

            protected override string CreateValue(UniqueID[] ids, object[,] correlArray)
            {
                correlArray = ExtensionMethods.ReIndexArray<object>(correlArray);
                StringBuilder sb = new StringBuilder();
                sb.Append($"{ids.Length},IM");
                sb.AppendLine();
                for (int field = 0; field < correlArray.GetLength(1); field++)
                {
                    //Add fields
                    sb.Append(ids[field].ID);
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
                return null;    //No names in IDs anymore
                //var ids = this.GetIDs();
                //return (from UniqueID uid in ids select uid.Name).ToArray<object>();
            }

            public static bool Validate(Excel.Range correlCell)      //validate that it is in fact a correlString
            {
                if (correlCell.NumberFormat == "\"Correl\";;;\"CORREL\"")
                    return true;
                else
                    return false;
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

            //private string ParseID(string id)
            //{
            //    string[] id_pieces = id.Split('|');         //split lines
            //    if (id_pieces.Length == 2)
            //        return id_pieces[1];                    //return the name portion of the ID
            //    else
            //        return null;                            //if malformed, return null
            //}

            public static CorrelationString_IM ConstructZeroString(string[] fields)
            {
                //Need to downcast csi 
                var csi = new CorrelationString(fields);
                return new CorrelationString_IM(csi.Value);
            }

            public static Data.CorrelationString_IM ConstructString(UniqueID[] ids, string sheet, Dictionary<Tuple<UniqueID, UniqueID>, double> correls = null)
            {
                Data.CorrelationString_IM correlationString = ConstructZeroString((from UniqueID id in ids select id.ID).ToArray());       //build zero string
                if (correls == null)
                    return correlationString;       //return zero string
                else
                {
                    Data.CorrelationMatrix matrix = new Data.CorrelationMatrix(correlationString);      //convert to zero matrix for modification
                    var matrixIDs = matrix.GetIDs();
                    foreach (UniqueID id1 in matrixIDs)
                    {
                        foreach (UniqueID id2 in matrixIDs)
                        {
                            if (correls.ContainsKey(new Tuple<UniqueID, UniqueID>(id1, id2)))
                            {
                                matrix.SetCorrelation(id1, id2, correls[new Tuple<UniqueID, UniqueID>(id1, id2)]);
                            }
                            if(correls.ContainsKey(new Tuple<UniqueID, UniqueID>(id2, id1)))
                            {
                                matrix.SetCorrelation(id2, id1, correls[new Tuple<UniqueID, UniqueID>(id2, id1)]);
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
