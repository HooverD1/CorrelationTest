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
            CostMatrix,
            CostTriple,
            PhasingMatrix,
            PhasingTriple,
            DurationMatrix,
            DurationTriple,
            Null
        }

        public class CorrelationString
        {
            public string Value { get; set; }
            public virtual string[] GetFields() { throw new Exception("Failed override"); }
            public virtual string[] GetIDs() { throw new Exception("Failed override"); }
            protected virtual string CreateValue(string parentID, object[] fields, object[,] correlArray) { throw new Exception("Failed override"); }
            protected virtual string CreateValue(string parentID, object[] ids, object[] fields, object[,] correlArray) { throw new Exception("Failed override"); }
            public virtual UniqueID GetParentID() { throw new Exception("Failed override"); }
            public virtual void Expand(Excel.Range xlSource) { throw new Exception("Failed override"); }
            public virtual void PrintToSheet(Excel.Range[] xlCells) { throw new Exception("Failed override"); }
            public virtual void PrintToSheet(Excel.Range xlCell) { PrintToSheet(new Excel.Range[] { xlCell }); }

            protected CorrelationString() { }
            public CorrelationString(string[] fields)     //creates zero string
            {
                this.Value = ExtensionMethods.CleanStringLinebreaks(CreateValue_Zero(fields));
            }

            public static string GetParentIDFromString(CorrelationString correlString_Object)
            {
                string correlString = Convert.ToString(correlString_Object.Value);
                correlString = ExtensionMethods.CleanStringLinebreaks(correlString);
                string[] lines = DelimitString(correlString);
                string[] header = lines[0].Split(',');
                return header[2];
            }

            public static string[] GetFieldsFromString(CorrelationString correlString_Object)
            {
                string correlString = Convert.ToString(correlString_Object.Value);
                correlString = ExtensionMethods.CleanStringLinebreaks(correlString);
                string[] lines = DelimitString(correlString);
                string[] fields = lines[1].Split(',');
                return fields;
            }
            
            public static string[] GetIDsFromString(object correlString_Object)
            {
                string correlString = Convert.ToString(correlString_Object);
                correlString = ExtensionMethods.CleanStringLinebreaks(correlString);
                string[] lines = DelimitString(correlString);
                string[] header = lines[0].Split(',');
                string[] ids = new string[header.Length - 3];
                for (int i = 3; i < header.Length; i++)
                    ids[i - 3] = header[i];
                return ids;
            }

            public string CreateValue_Zero(string[] fields, double defaultValue = 0) //create a zero correlstring from very generic params
            {
                StringBuilder sb = new StringBuilder();
                sb.Append($"{fields.Length},CM");
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

            public static string[] DelimitString(string correlStringValue)
            {
                string[] correlLines = correlStringValue.Split('&');         //split lines
                return correlLines;
            }

            public virtual object[,] GetMatrix()
            {       //returning 2,2 instead of 3,3
                string myValue = ExtensionMethods.CleanStringLinebreaks(this.Value);
                string[] fieldString1 = myValue.Split('&');          //broken by line
                string[] fieldString = new string[fieldString1.Length - 2];
                for (int i = 2; i < fieldString1.Length; i++) { fieldString[i - 2] = fieldString1[i]; }  //dump the header and fields
                object[,] matrix = new object[fieldString.Length + 1, fieldString.Length + 1];

                for (int row = 0; row < fieldString.Length + 1; row++)
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
                            if (Double.TryParse(values[(col - row) - 1], out double conversion))
                            {
                                matrix[row, col] = conversion;
                            }
                        }


                        else  //col < row
                            matrix[row, col] = null;
                    }
                }
                return matrix;
                throw new Exception("Failed override");
            }

            public virtual string GetCorrelType()
            {
                string[] lines = DelimitString(this.Value);
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

            public int GetNumberOfSubs()
            {
                string[] lines = DelimitString(this.Value);
                return Convert.ToInt32(lines[0].Split(',')[0]);
            }

            #region CorrelString Factory
            private static CorrelStringType ParseCorrelType(string correlStringValue)
            {
                correlStringValue = ExtensionMethods.CleanStringLinebreaks(correlStringValue);
                string[] splitValues = correlStringValue.Split('&')[0].Split(',');
                // # Periods | Type Char
                string correlTypeStr = splitValues[1];
                switch (correlTypeStr)
                {
                    case "CM":
                        return CorrelStringType.CostMatrix;
                    case "CT":
                        return CorrelStringType.CostTriple;
                    case "PM":
                        return CorrelStringType.PhasingMatrix;
                    case "PT":
                        return CorrelStringType.PhasingTriple;
                    case "DT":
                        return CorrelStringType.DurationTriple;
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
                    case CorrelStringType.CostTriple:
                        return ParseInputsTriple(stringValues);
                    case CorrelStringType.CostMatrix:
                        return ParseInputsMatrix(stringValues);
                    case CorrelStringType.DurationTriple:
                        return ParseDurationTriple(stringValues);
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
            private static Dictionary<string, object> ParseInputsTriple(string[][] values)
            {
                Dictionary<string, object> dict = new Dictionary<string, object>();
                dict.Add("Children", values[0][0]);
                dict.Add("Parent_ID", values[1][0]);
                dict.Add("Triple", $"{values[2][0]},{values[2][1]},{values[2][2]}");
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
            private static Dictionary<string, object> ParseDurationTriple(string[][] values)
            {
                throw new NotImplementedException();
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

            public static bool Validate(string correlStringValue)
            {
                //act as a switch for sending a string to its proper subclass validation
                return true;
            }

            public static CorrelationString ConstructFromCorrelSheet()
            {
                //Need to get the values from the matrix
                //Need to get the string type from the header
                //Need to get the parentID from the header

                //Get the type of correlation sheet
                Sheets.CorrelationSheet cSheet = Sheets.CorrelationSheet.Construct();
                return cSheet.CorrelString;                
                //CREATE VALUE: string parentID, object[] fields, object[,] correlArray
            }

            public static CorrelationString ConstructFromCorrelationSheet(Sheets.CorrelationSheet correlSheet)
            {
                string sheetTag = Convert.ToString(correlSheet.xlSheet.Cells[1, 1].value);
                switch(sheetTag)
                {
                    case "$CORRELATION_CT":
                        return new CorrelationString_CT((Sheets.CorrelationSheet_Cost)correlSheet);
                    case "$CORRELATION_CM":
                        return new CorrelationString_CM((Sheets.CorrelationSheet_Cost)correlSheet);
                    case "$CORRELATION_PT":
                        return new CorrelationString_PT((Sheets.CorrelationSheet_Phasing)correlSheet);
                    case "$CORRELATION_PM":
                        return new CorrelationString_PM((Sheets.CorrelationSheet_Phasing)correlSheet);
                    case "$CORRELATION_DT":
                        return new CorrelationString_DT((Sheets.CorrelationSheet_Duration)correlSheet);
                    case "$CORRELATION_DM":
                        return new CorrelationString_DM((Sheets.CorrelationSheet_Duration)correlSheet);
                    default:
                        throw new Exception("Malformed correlation string");
                }
            }

            public static CorrelationString ConstructFromStringValue(string correlStringValue)
            {
                string[] lines = DelimitString(correlStringValue);
                string[] header = lines[0].Split(',');
                switch (header[1])
                {
                    case "CT":
                        return new CorrelationString_CT(correlStringValue);
                    case "CM":
                        return new CorrelationString_CM(correlStringValue);
                    case "PT":
                        return new CorrelationString_PT(correlStringValue);
                    case "PM":
                        return new CorrelationString_PM(correlStringValue);
                    case "DT":
                        return new CorrelationString_DT(correlStringValue);
                    case "DM":
                        return new CorrelationString_DM(correlStringValue);
                    default:
                        throw new Exception("Malformed correlation string");
                }
            }

            public static CorrelationString ConstructDefaultFromCostSheet(IHasSubs item, CorrelStringType csType)        //Construct default correlation string for estimate
            {
                switch (csType)
                {
                    case CorrelStringType.PhasingTriple:
                        Triple pt = new Triple(item.uID.ID, "0,0,0");
                        string[] start_dates = ((IHasPhasingSubs)item).Periods.Select(x => x.Start_Date).ToArray();
                        return new Data.CorrelationString_PT(pt, start_dates, item.uID.ID);
                    case CorrelStringType.PhasingMatrix:
                        IEnumerable<string> start_dates2 = from Period prd in ((IHasPhasingSubs)item).Periods select prd.pID.PeriodTag.ToString();
                        return CorrelationString_PM.ConstructZeroString(start_dates2.ToArray());
                    case CorrelStringType.CostTriple:
                        if (((IHasCostSubs)item).SubEstimates.Count < 2)
                            return null;
                        Triple it = new Triple(item.uID.ID, "0,0,0");
                        IEnumerable<string> fields = from ISub sub in ((IHasCostSubs)item).SubEstimates select sub.Name;        //need to print names, but get them from IDs?
                        return new Data.CorrelationString_CT(fields.ToArray(), it, item.uID.ID, ((IHasCostSubs)item).SubEstimates.Select(x => x.uID.ID).ToArray());
                    case CorrelStringType.CostMatrix:
                        if (((IHasCostSubs)item).SubEstimates.Count < 2)
                            return null;
                        IEnumerable<string> fields2 = from Estimate_Item sub in item.ContainingSheetObject.GetSubEstimates(item.xlRow) select sub.Name;
                        return CorrelationString_CM.ConstructZeroString(fields2.ToArray());
                    case CorrelStringType.DurationMatrix:
                        throw new NotImplementedException();
                    case CorrelStringType.DurationTriple:
                        if (((IHasDurationSubs)item).SubEstimates.Count < 2)
                            return null;
                        Triple it2 = new Triple(item.uID.ID, "0,0,0");
                        IEnumerable<string> fields3 = from ISub sub in ((IHasDurationSubs)item).SubEstimates select sub.Name;
                        return new Data.CorrelationString_DT(fields3.ToArray(), it2, item.uID.ID, ((IHasDurationSubs)item).SubEstimates.Select(x => x.uID.ID).ToArray());
                    default:
                        throw new Exception("Cannot construct CorrelationString");
                }
            }

            public static string ConstructStringFromRange(IEnumerable<Excel.Range> stringRanges)
            {
                //Pull the fragments of a correlation string off the sheet and recombine into one string
                StringBuilder sb = new StringBuilder();
                foreach(Excel.Range strRange in stringRanges)
                {
                    if(strRange.Value != null)
                    {
                        sb.Append(Convert.ToString(strRange.Value));
                        sb.Append("&");
                    }
                }
                sb.Remove(sb.Length - 1, 1);    //remove the final &
                return sb.ToString();
            }
            #endregion

        }
    }
    
}
