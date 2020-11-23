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

            public CorrelationString(Excel.Range xlRange) : this(GetCorrelArrayFromRange(xlRange), GetFieldsFromRange(xlRange)) { }
            
            public CorrelationString(string correlString)
            {
                this.Value = correlString;
            }
            public CorrelationString(object[,] correlArray, object[] fields)
            {
                this.Value = CreateValue(correlArray, fields);               
            }

            private CorrelationString(object[] fields)     //create 0 string (independence)
            {
                int fieldCount = fields.Count();
                object[,] correlArray = new object[fieldCount, fieldCount];
                for(int row = 0; row < fieldCount; row++)
                {
                    for(int col = 0; col < fieldCount; col++)
                    {
                        correlArray[row, col] = 0;
                    }
                }
                this.Value = CreateValue(correlArray, fields);
            }
            
            private static object[] GetFieldsFromRange(Excel.Range correlRange)
            {
                object[,] tempArray = correlRange.Resize[1, correlRange.Columns.Count].Value;
                object[] returnArray = new object[tempArray.GetLength(1)];
                for(int i = 0; i < tempArray.GetLength(1); i++)
                {
                    returnArray[i] = tempArray[1, i+1];
                }
                return returnArray;
            }
            private static object[,] GetCorrelArrayFromRange(Excel.Range correlRange)
            {
                return correlRange.Offset[1, 0].Resize[correlRange.Rows.Count - 1, correlRange.Columns.Count].Value;
            }
            private string CreateValue(object[,] correlArray, object[] fields)
            {
                correlArray = ExtensionMethods.ReIndexArray<object>(correlArray);
                fields = ExtensionMethods.ReIndexArray<object>(fields);
                StringBuilder sb = new StringBuilder();
                for (int field = 0; field < correlArray.GetLength(1); field++)
                {
                    //Add fields
                    sb.Append(fields[field].ToString());
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
            public object[] GetFields()
            {
                string correlString = this.Value.Replace("\r\n", "&");  //simplify delimiter
                string[] correlLines = correlString.Split('&');         //split lines
                string[] fields = correlLines[0].Split(',');            //get fields (first line) and delimit
                return Array.ConvertAll<string, object>(fields, new Converter<string, object>(x => (object)x));
            }
            public object[,] GetMatrix()
            {
                string myValue = this.Value.Replace("\r\n", "&");
                string[] fieldString = myValue.Split('&');          //broken by line
                
                object[,] matrix = new object[fieldString.Length-1, fieldString.Length-1];
                
                for (int row = 0; row < fieldString.Length-1; row++)
                {
                    string[] values;
                    values = fieldString[row + 1].Split(',');       //broken by entry

                    for (int col = 0; col < fieldString.Length-1; col++)
                    {
                        if (col > row)
                            matrix[row, col] = values[(col-row) - 1];
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

            public static Data.CorrelationString ConstructString(object[] fields, Estimate parent = null)
            {
                Data.CorrelationString correlationString = new CorrelationString(fields);       //zero string
                if (parent == null)
                    return correlationString;
                else
                {
                    //if you send in a parent estimate, pull its children's NonZero items
                    //do I need to construct the children here?
                    Data.CorrelationMatrix matrix = new Data.CorrelationMatrix(correlationString);      //zero matrix
                    foreach (Estimate sub in parent.SubEstimates)
                    {
                        foreach(KeyValuePair<string, double> pair in sub.CorrelPairs)
                        {
                            matrix.SetCorrelation(sub.ID, pair.Key, pair.Value);            //override zeroes
                        }
                    }
                }
                return correlationString;
            }

            public static void ExpandCorrel(Excel.Range selection)
            {
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
