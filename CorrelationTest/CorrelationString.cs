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
        public class CorrelationString
        {
            public string Value { get; set; }

            public CorrelationString(Excel.Range xlRange) : this(GetCorrelArrayFromRange(xlRange), GetIDsFromRange(xlRange), xlRange.Worksheet.Name) { }
            
            public CorrelationString(string correlString)
            {
                this.Value = correlString;
            }
            public CorrelationString(object[,] correlArray, UniqueID[] ids, string sheet)
            {
                this.Value = CreateValue(correlArray, ids);               
            }

            private CorrelationString(UniqueID[] ids, string sheet)     //create 0 string (independence)
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
                this.Value = CreateValue(correlArray, ids);
            }

            public CorrelationString(Data.CorrelationMatrix matrix)
            {
                this.Value = CreateValue(matrix.GetMatrix(), matrix.GetIDs());
            }

            private object[] BuildIDsFromFields(object[] fields, string sheet)
            {
                object[] ids = new object[fields.Length];
                for (int i = 0; i < fields.Length; i++)
                    ids[i] = $"{sheet}|{fields[i]}";
                return ids;
            }
            
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
                    sb.Append(parentEstimate.SubEstimates[sub].ID);
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

            private string CreateValue(object[,] correlArray, UniqueID[] ids)
            {
                correlArray = ExtensionMethods.ReIndexArray<object>(correlArray);
                StringBuilder sb = new StringBuilder();
                
                for (int field = 0; field < correlArray.GetLength(1); field++)
                {
                    //Add fields
                    sb.Append(ids[field].Value);
                    if(field < correlArray.GetLength(1)-1)
                        sb.Append(",");
                }
                sb.AppendLine();
                for (int row = 0; row < correlArray.GetLength(0); row++)
                {
                    for (int col = row+1; col < correlArray.GetLength(1); col++)
                    {
                        sb.Append(correlArray[row, col]);
                        if(col < correlArray.GetLength(1)-1)
                            sb.Append(",");
                    }
                    if (row < correlArray.GetLength(0) - 1)
                        sb.AppendLine();
                }
                return sb.ToString();
            }



            public string CreateArray(string correlString)
            {
                throw new NotImplementedException();
            }

            public string[] GetFields(double[,] correlArray)
            {
                throw new NotImplementedException();
            }
            public double[,] UpdateFields(double[,] getFields)
            {
                throw new NotImplementedException();
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
            public UniqueID[] GetIDs()
            {
                string correlString = this.Value.Replace("\r\n", "&");  //simplify delimiter
                       correlString = correlString.Replace("\n", "&");  //simplify delimiter
                string[] correlLines = correlString.Split('&');         //split lines
                string[] id_strings = correlLines[0].Split(',');            //get fields (first line) and delimit
                UniqueID[] returnIDs = id_strings.Select(x => new UniqueID(x)).ToArray();
                if (returnIDs.Distinct().Count() == returnIDs.Count())
                    return returnIDs;
                else
                    return UniqueID.AutoFixUniqueIDs(returnIDs);
                //return Array.ConvertAll<string, object>(ids, new Converter<string, object>(x => (object)x));
            }
            public object[] GetFields()
            {
                List<string> returnList = new List<string>();
                foreach(UniqueID id in GetIDs())
                {
                    returnList.Add(id.FieldName);
                }
                return returnList.ToArray<object>();
            }
            //private string ParseID(string id)
            //{
            //    string[] id_pieces = id.Split('|');         //split lines
            //    if (id_pieces.Length == 2)
            //        return id_pieces[1];                    //return the name portion of the ID
            //    else
            //        return null;                            //if malformed, return null
            //}
            public object[,] GetMatrix()
            {
                string myValue = this.Value.Replace("\r\n", "&");
                       myValue = myValue.Replace("\n", "&");
                string[] fieldString = myValue.Split('&');          //broken by line
                
                object[,] matrix = new object[fieldString.Length-1, fieldString.Length-1];
                
                for (int row = 0; row < fieldString.Length-1; row++)
                {
                    string[] values;
                    values = fieldString[row + 1].Split(',');       //broken by entry

                    for (int col = 0; col < fieldString.Length-1; col++)
                    {
                        if (col > row)
                            matrix[row, col] = Convert.ToDouble(values[(col-row) - 1]);
                        else if (col == row)
                            matrix[row, col] = 1;
                        else  //col < row
                            matrix[row, col] = null;
                    }
                }
                return matrix;
            }
            public void PrintToSheet(Excel.Range xlOrigin)
            {
                xlOrigin.Value = this.Value;
                xlOrigin.WrapText = false;
                xlOrigin.NumberFormat = "\"Correl\";;;\"CORREL\"";
                xlOrigin.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            }

            public static Data.CorrelationString ConstructString(UniqueID[] ids, string sheet, Dictionary<Tuple<UniqueID, UniqueID>, double> correls = null)
            {
                Data.CorrelationString correlationString = new CorrelationString(ids, sheet);       //build zero string
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
                    return new Data.CorrelationString(matrix);      //return modified zero matrix as correl string
                }
            }

            public static void ExpandCorrel(Excel.Range selection)
            {
                //Check if in edit mode

                //Verify that it's a correl string
                bool valid = CorrelationString.Validate(selection);
                if (valid)
                {
                    //Construct the estimate
                    IEstimate tempEstimate = EstimateFactory.Construct(selection);
                    //construct the correlString
                    Data.CorrelationString correlStringObj = new Data.CorrelationString(Convert.ToString(tempEstimate.xlCorrelCell.Value));
                    //construct the correlSheet
                    Sheets.CorrelationSheet correlSheet = new Sheets.CorrelationSheet(correlStringObj, selection, new Data.CorrelSheetSpecs());
                    //print the correlSheet
                    correlSheet.PrintToSheet();
                }
            }
        }
    }
}
