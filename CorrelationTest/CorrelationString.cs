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
            InputsTriple,
            PhasingMatrix,
            PhasingTriple,
            DurationMatrix,
            DurationTriple
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
                this.Value = ExtensionMethods.CleanStringLinebreaks(CreateValue_Zero(fields));
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
                string correlString = this.Value;
                string[] correlLines = correlString.Split('&');         //split lines
                return correlLines;
            }

            public virtual object[,] GetMatrix()
            {       //returning 2,2 instead of 3,3
                string myValue = ExtensionMethods.CleanStringLinebreaks(this.Value);
                string[] fieldString1 = myValue.Split('&');          //broken by line
                string[] fieldString = new string[fieldString1.Length - 2];
                for(int i = 2; i < fieldString1.Length; i++) { fieldString[i - 2] = fieldString1[i]; }  //dump the header and fields
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

            public virtual string GetCorrelType()
            {
                string[] lines = DelimitString();
                return lines[0].Split(',')[1];
            }

            public virtual bool ValidateAgainstMatrix(object[] outsideFields)
            {
                var localFields = this.GetFields();
                if (localFields.Count() != outsideFields.Count())
                {
                    return false;
                }
                foreach (object field in localFields)
                {
                    if (!outsideFields.Contains<object>(field))
                        return false;
                }
                return true;
            }

            public int GetNumberOfPeriods()
            {
                string[] lines = DelimitString();
                return Convert.ToInt32(lines[0].Split(',')[0]);
            }

            protected virtual string CreateValue(UniqueID[] ids, object[,] correlArray)
            {
                return null;

                //correlArray = ExtensionMethods.ReIndexArray<object>(correlArray);
                //StringBuilder sb = new StringBuilder();
                //sb.AppendLine();
                //sb.Append($"{ids.Length},");
                //sb.AppendLine();
                //for (int field = 0; field < correlArray.GetLength(1); field++)
                //{
                //    //Add fields
                //    sb.Append(ids[field].ID);
                //    if (field < correlArray.GetLength(1) - 1)
                //        sb.Append(",");
                //}
                //sb.AppendLine();
                //for (int row = 0; row < correlArray.GetLength(0); row++)
                //{
                //    for (int col = row + 1; col < correlArray.GetLength(1); col++)
                //    {
                //        sb.Append(correlArray[row, col]);
                //        if (col < correlArray.GetLength(1) - 1)
                //            sb.Append(",");
                //    }
                //    if (row < correlArray.GetLength(0) - 2)
                //        sb.AppendLine();
                //}
                //return sb.ToString();
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

            public static bool Validate(string correlString)
            {
                string[][] ss = SplitString(correlString);
                if (ss.GetLength(0) < 3)        //header, field, and at least one value row
                    return false;
                if (ss[0].Length != 2)          //header has two components
                    return false;
                foreach(string s in ss[2])      //values can all be resolved to doubles
                {
                    bool pass = false;
                    pass = double.TryParse(s, out double result);
                    if (!pass)
                        return false;
                }
                return true;
            }

            public static CorrelationString Construct(Estimate est, CorrelStringType csType)        //Construct default correlation string for estimate
            {
                switch (csType)
                {
                    case CorrelStringType.PhasingTriple:
                        if(est.xlCorrelCell_Periods.Value == null)
                        {
                            Triple pt = new Triple(est.uID.ID, "0,0,0");
                            return new Data.CorrelationString_PT(pt, est.Periods.Length, est.uID.ID);
                        }
                        else
                        {
                            return Construct(est.xlCorrelCell_Periods.Value);
                        }                        
                    case CorrelStringType.PhasingMatrix:
                        if(est.xlCorrelCell_Periods.Value == null)
                        {
                            IEnumerable<string> start_dates = from Period prd in est.Periods select prd.pID.PeriodTag.ToString();
                            return CorrelationString_PM.ConstructZeroString(start_dates.ToArray());
                        }
                        else
                        {
                            return Construct(est.xlCorrelCell_Periods.Value);
                        }
                    case CorrelStringType.InputsTriple:
                        if (est.xlCorrelCell_Inputs.Value == null)
                        {
                            if (est.SubEstimates.Count < 2)
                                return null;
                            Triple it = new Triple(est.uID.ID, "0,0,0");
                            return new Data.CorrelationString_IT(it, est.SubEstimates.Count, est.uID.ID);
                        }
                        else
                        {
                            return Construct(est.xlCorrelCell_Inputs.Value);
                        }
                    case CorrelStringType.InputsMatrix:
                        if(est.xlCorrelCell_Inputs.Value == null)
                        {
                            if (est.SubEstimates.Count < 2)
                                return null;
                            IEnumerable<string> fields = from Estimate sub in est.ContainingSheetObject.GetSubEstimates(est.xlRow) select sub.Name;
                            return CorrelationString_IM.ConstructZeroString(fields.ToArray());
                        }
                        else
                        {
                            return Construct(est.xlCorrelCell_Inputs.Value);
                        }                        
                    case CorrelStringType.DurationMatrix:
                        throw new NotImplementedException();
                    case CorrelStringType.DurationTriple:
                        throw new NotImplementedException();
                    default:
                        throw new Exception("Cannot construct CorrelationString");
                }
            }

            public static CorrelationString Construct(string correlStringValue)     //construct a variety of CorrelationStrings from the string
            {
                //validate that it is a valid correlation string
                if (!Validate(correlStringValue))
                    throw new Exception("Invalid correlation string.");
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
                        Triple pt = new Triple((string)parameters["Parent_ID"], (string)parameters["Triple"]);
                        return new Data.CorrelationString_PT(pt, Convert.ToInt32(parameters["Periods"]), values[1][0].ToString());
                    case CorrelStringType.PhasingMatrix:
                        return new CorrelationString_PM(correlStringValue);
                    case CorrelStringType.InputsTriple:
                        return new CorrelationString_IT(correlStringValue);
                    case CorrelStringType.InputsMatrix:
                        return new CorrelationString_IM(correlStringValue);
                    case CorrelStringType.DurationTriple:
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
