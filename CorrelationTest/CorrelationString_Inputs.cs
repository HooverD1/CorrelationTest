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
        public class CorrelationString_Inputs : CorrelationString
        {
            public CorrelationString_Inputs(Excel.Range xlRange) : this(GetCorrelArrayFromRange(xlRange), GetIDsFromRange(xlRange)) { }
            
            

            public CorrelationString_Inputs(string correlString)
            {
                this.Value = correlString;
            }
            public CorrelationString_Inputs(object[,] correlArray, UniqueID[] ids)
            {
                this.Value = CreateValue(ids, correlArray);               
            }

            private CorrelationString_Inputs(UniqueID[] ids, string sheet)     //create 0 string (independence)
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
                this.Value = CreateValue(ids, correlArray);
            }

            public CorrelationString_Inputs(Data.CorrelationMatrix matrix)
            {
                this.Value = CreateValue(matrix.GetIDs(), matrix.GetMatrix());
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
                var specs = new CorrelSheetSpecs();
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
                StringBuilder sb = new StringBuilder();
                int fields = parentEstimate.SubEstimates.Count;
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

            public string CreateArray(string correlString)
            {
                throw new NotImplementedException();
            }

            public override object[] GetFields()
            {
                var ids = this.GetIDs();
                return (from UniqueID uid in ids select uid.Name).ToArray<object>();
            }

            public bool ValidateAgainstMatrix(object[] outsideFields)
            {
                var localFields = this.GetFields();
                if(localFields.Count() != outsideFields.Count())
                {
                    return false;
                }
                foreach(object field in localFields)
                {
                    if (!outsideFields.Contains<object>(field))
                        return false;
                }
                return true;
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
                string[] id_strings = correlLines[0].Split(',');            //get fields (first line) and delimit
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

            public static CorrelationString_Inputs CreateZeroString(string[] fields)
            {
                //Need to downcast csi 
                var csi = new CorrelationString(fields);
                return new CorrelationString_Inputs(csi.Value);
            }

            public static Data.CorrelationString_Inputs ConstructString(UniqueID[] ids, string sheet, Dictionary<Tuple<UniqueID, UniqueID>, double> correls = null)
            {
                Data.CorrelationString_Inputs correlationString = CreateZeroString((from UniqueID id in ids select id.ID).ToArray());       //build zero string
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
                    return new Data.CorrelationString_Inputs(matrix);      //return modified zero matrix as correl string
                }
            }

            public static void ExpandCorrel(Excel.Range selection)
            {
                //Verify that it's a correl string
                bool valid = CorrelationString_Inputs.Validate(selection);
                if (valid)
                {
                    //Construct the estimate
                    IEstimate tempEstimate = EstimateFactory.Construct(selection);
                    //construct the correlString
                    Data.CorrelationString_Inputs correlStringObj = new Data.CorrelationString_Inputs(Convert.ToString(tempEstimate.xlCorrelCell.Value));
                    //construct the correlSheet
                    Sheets.CorrelationSheet correlSheet = new Sheets.CorrelationSheet(correlStringObj, selection, new Data.CorrelSheetSpecs());
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
                for(int i=0; i < newIDs.Length; i++)
                {
                    sb.Append(newIDs[i].Name);
                    if (i < newIDs.Length - 1)
                        sb.Append(",");
                }
                sb.AppendLine();
                for(int j=1;j<correlLines.Length;j++)
                {
                    sb.Append(correlLines[j]);
                }
                this.Value = sb.ToString();
            }


        }
    }
}
