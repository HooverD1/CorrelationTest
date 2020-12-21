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
        public enum CorrelStringType
        {
            InputsMatrix,
            PhasingMatrix,
            PhasingTriple,
            DurationMatrix
        }

        public class CorrelationString
        {
            public string Value { get; set; }
            public virtual object[] GetFields() { return null; }
            public virtual UniqueID[] GetIDs() { return null; }

            public virtual void PrintToSheet(Excel.Range xlCell) { }
            protected CorrelationString() { }

            public CorrelationString(string[] fields)     //creates zero string
            {
                this.Value = CreateValue_Zero(fields);
            }



            public string CreateValue_Zero(string[] fields, double defaultValue = 0) //create a zero correlstring from very generic params
            {
                StringBuilder sb = new StringBuilder();
                sb.Append($"{fields.Length},IM");
                sb.AppendLine();
                for (int i = 0; i < fields.Length; i++)
                {
                    sb.Append(fields[i]);
                    if (i < fields.Length - 1)
                        sb.Append(",");
                }
                sb.AppendLine();
                for (int row = 0; row < fields.Length-1; row++)
                {
                    for (int col = row+1; col < fields.Length; col++)
                    {
                        if(col > row)
                        {
                            sb.Append(defaultValue.ToString());
                        }
                        else
                        {
                            continue;
                        }
                        if (col < fields.Length-1)
                            sb.Append(",");
                    }
                    if (row < fields.Length - 2)
                        sb.AppendLine();
                }
                return sb.ToString();
            }

            protected string[] DelimitString()
            {
                string correlString = ExtensionMethods.CleanStringLinebreaks(this.Value);
                string[] correlLines = correlString.Split('&');         //split lines
                return correlLines;
            }

            public object[,] GetMatrix()
            {       //returning 2,2 instead of 3,3
                string myValue = ExtensionMethods.CleanStringLinebreaks(this.Value);
                string[] fieldString1 = myValue.Split('&');          //broken by line
                string[] fieldString = new string[fieldString1.Length - 1];
                for(int i = 1; i < fieldString1.Length; i++) { fieldString[i - 1] = fieldString1[i]; }
                object[,] matrix = new object[fieldString.Length+1, fieldString.Length+1];

                for (int row = 0; row < fieldString.Length+1; row++)
                {
                    string[] values;
                    if (row < fieldString.Length)
                        values = fieldString[row].Split(',');       //broken by entry
                    else
                        values = null;

                    for (int col = fieldString.Length; col >= 0; col--)
                    {
                        if (col == row)
                            matrix[row, col] = 1;
                        else if (col > row)
                        {
                            if(Double.TryParse(values[(col - row) - 1], out double conversion))
                            {
                                matrix[row, col] = conversion;
                            }
                        }
                            

                        else  //col < row
                            matrix[row, col] = null;
                    }
                }
                return matrix;
            }

            protected virtual string CreateValue(UniqueID[] ids, object[,] correlArray)
            {
                correlArray = ExtensionMethods.ReIndexArray<object>(correlArray);
                StringBuilder sb = new StringBuilder();

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

            public virtual UniqueID GetParentID() { return null; }

            public virtual void Expand(Excel.Range xlSource) { }

            #region CorrelString Factory
            private static CorrelStringType ParseCorrelType(string correlStringValue)
            {
                correlStringValue = ExtensionMethods.CleanStringLinebreaks(correlStringValue);
                string[] splitValues = correlStringValue.Split('&')[0].Split(',');
                // # Periods | Type Char
                string correlTypeStr = splitValues[1];
                switch (correlTypeStr)
                {
                    case "IM":
                        return CorrelStringType.InputsMatrix;
                    case "PM":
                        return CorrelStringType.PhasingMatrix;
                    case "PT":
                        return CorrelStringType.PhasingTriple;
                    case "DM":
                        return CorrelStringType.DurationMatrix;
                    default:
                        throw new Exception("Malformed Correlation String Header");
                }
            }

            private static Dictionary<string, object> ParseStringValue(string[][] stringValues, CorrelStringType csType)
            {
                Dictionary<string, object> stringDictionary = new Dictionary<string, object>();
                switch (csType)
                {
                    case CorrelStringType.PhasingTriple:
                        return ParsePhasingTriple(stringValues);
                    case CorrelStringType.PhasingMatrix:
                        return ParsePhasingMatrix(stringValues);
                    case CorrelStringType.InputsMatrix:
                        return ParseInputsMatrix(stringValues);
                    case CorrelStringType.DurationMatrix:
                        return ParseDurationMatrix(stringValues);
                    default:
                        throw new Exception("Malformed CorrelStringType");
                }
            }

            private static Dictionary<string, object> ParsePhasingTriple(string[][] values)
            {
                Dictionary<string, object> dict = new Dictionary<string, object>();
                dict.Add("Periods", values[0][0]);
                dict.Add("Parent_ID", values[1][0]);
                //object[,] tripleArray = ExtensionMethods.GetSubArray(values, 2);
                dict.Add("Triple", $"{values[2][0]},{values[2][1]},{values[2][2]}");
                return dict;
            }
            private static Dictionary<string, object> ParsePhasingMatrix(string[][] values)
            {
                Dictionary<string, object> dict = new Dictionary<string, object>();
                dict.Add("Periods", values[0][0]);
                dict.Add("IDs", values[1]);
                dict.Add("Matrix", ExtensionMethods.GetSubArray(values, 2));
                return dict;
            }
            private static Dictionary<string, object> ParseInputsMatrix(string[][] values)
            {
                Dictionary<string, object> dict = new Dictionary<string, object>();
                dict.Add("Children", values[0][0]);
                dict.Add("IDs", values[1]);
                dict.Add("Matrix", ExtensionMethods.GetSubArray(values, 2));
                return dict;
            }
            private static Dictionary<string, object> ParseDurationMatrix(string[][] values)
            {
                Dictionary<string, object> dict = new Dictionary<string, object>();
                dict.Add("Children", values[0][0]);
                dict.Add("IDs", values[1]);
                dict.Add("Matrix", ExtensionMethods.GetSubArray(values, 2));
                return dict;
            }

            private static string[][] SplitString(string correlStringValue)
            {
                Dictionary<string, object> dict = new Dictionary<string, object>();
                string[] lines = correlStringValue.Split('&');
                string[][] values = new string[lines.Length][];
                for (int i = 0; i < lines.Length; i++)
                {
                    values[i] = lines[i].Split(',');
                }
                return values;
            }

            public static CorrelationString Construct(string correlStringValue)     //construct a variety of CorrelationStrings from the string
            {
                CorrelStringType csType = ParseCorrelType(correlStringValue);
                correlStringValue = ExtensionMethods.CleanStringLinebreaks(correlStringValue);
                string[][] values = SplitString(correlStringValue);
                Dictionary<string, object> parameters = ParseStringValue(values, csType);
                //parse string values
                //return a dictionary
                //use that to build the object
                switch (csType)
                {
                    case CorrelStringType.PhasingTriple:
                        PhasingTriple pt = new PhasingTriple((string)parameters["Parent_ID"], (string)parameters["Triple"]);
                        return pt.GetCorrelationString(Convert.ToInt32(parameters["Periods"]), values[1][0].ToString());
                    case CorrelStringType.PhasingMatrix:
                        return new CorrelationString_Periods(correlStringValue);
                    case CorrelStringType.InputsMatrix:
                        throw new NotImplementedException();
                    case CorrelStringType.DurationMatrix:
                        throw new NotImplementedException();
                    default:
                        throw new Exception("Cannot construct CorrelationString");
                }
            }

            #endregion

        }
    }
    
}
