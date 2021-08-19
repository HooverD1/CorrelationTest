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
            CostPair,
            PhasingMatrix,
            PhasingPair,
            DurationMatrix,
            DurationPair,
            Null
        }

        public class CorrelationString
        {
            public string Value { get; set; }
            public virtual string[] GetIDs() { throw new Exception("Failed override"); }
            protected virtual string CreateValue(string parentID, object[] fields, object[,] correlArray) { throw new Exception("Failed override"); }
            protected virtual string CreateValue(string parentID, object[] ids, object[] fields, object[,] correlArray) { throw new Exception("Failed override"); }
            public virtual UniqueID GetParentID() { throw new Exception("Failed override"); }
            public virtual void PrintToSheet(Excel.Range[] xlCells) { throw new Exception("Failed override"); }
            public virtual void PrintToSheet(Excel.Range xlCell) { PrintToSheet(new Excel.Range[] { xlCell }); }

            protected CorrelationString() { }
            public CorrelationString(string[] fields)     //creates zero string
            {
                this.Value = ExtensionMethods.CleanStringLinebreaks(CreateValue_Zero(fields));
            }

            public string GetHeader()
            {
                string correlString = ExtensionMethods.CleanStringLinebreaks(this.Value);
                string[] lines = DelimitString(correlString);
                return lines[0];
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

            public virtual string[,] GetMatrix_Formulas(Sheets.CorrelationSheet CorrelSheet) { throw new Exception("Failed override"); }

            //Gets the double matrix for a matrix spec'ed correl string
            public virtual double[,] GetMatrix_Doubles() { throw new Exception("Failed override"); }

            public virtual string[,] GetMatrix_Values()     //What is this for? Displaying matrix specs?
            {
                string myValue = ExtensionMethods.CleanStringLinebreaks(this.Value);
                string[] lines = DelimitString(myValue);
                string[] header = lines[0].Split(',');
                int length = Convert.ToInt32(header[0]);
                string[,] matrix = new string[length, length];

                for (int row = 0; row < length; row++)
                {
                    string[] values;
                    if (row + 1 < length)
                        values = lines[row + 1].Split(',');       //broken by entry
                    else
                        values = null;

                    for (int col = row; col < length; col++)
                    {
                        if (col == row)
                            matrix[row, col] = "1";
                        else if (col > row && values != null)
                        {
                            if (Double.TryParse(values[(col - row) - 1], out double conversion))
                            {
                                matrix[row, col] = conversion.ToString();
                            }
                        }

                    }
                }

                //Fill in lower triangular
                for (int row = 0; row < length; row++)
                {
                    for (int col = 0; col < row; col++)
                    {
                        matrix[row, col] = $"=OFFSET(INDIRECT(ADDRESS(ROW(),COLUMN(),4,1)),-{row-col},{row-col})";
                    }
                }

                return matrix;
            }

            public SheetType GetCorrelType()
            {
                string[] lines = DelimitString(this.Value);
                switch (lines[0].Split(',')[1])
                {
                    case "CM":
                        return SheetType.Correlation_CM;
                    case "CP":
                        return SheetType.Correlation_CP;
                    case "PM":
                        return SheetType.Correlation_PM;
                    case "PP":
                        return SheetType.Correlation_PP;
                    case "DM":
                        return SheetType.Correlation_DM;
                    case "DP":
                        return SheetType.Correlation_DP;
                    default:
                        throw new Exception("Unknown correl type");
                }
            }

            public virtual bool ValidateAgainstMatrix(object[] outsideIDs)
            {
                var localIDs = this.GetIDs();
                if (localIDs.Count() != outsideIDs.Count())
                {
                    return false;
                }
                foreach (object field in localIDs)
                {
                    if (!outsideIDs.Contains<object>(field))
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

            public static CorrelationString_CP ConstructFromParentItem_Cost(IHasCostCorrelations ParentItem)
            {
                StringBuilder header = new StringBuilder();
                StringBuilder values = new StringBuilder();

                header.Append(ParentItem.SubEstimates.Count);
                header.Append(",");

                header.Append("CP");
                header.Append(",");

                header.Append(ParentItem.uID.ID);
                for(int i = 0; i < ParentItem.SubEstimates.Count; i++)
                {
                    header.Append(",");
                    header.Append(ParentItem.SubEstimates[i].uID.ID);
                }
                for(int i = 0; i < ParentItem.SubEstimates.Count; i++)
                {
                    values.Append("&0,0");
                }
                    
                return new CorrelationString_CP($"{header}{values}");
            }

            public static CorrelationString_PP ConstructFromParentItem_Phasing(IHasPhasingCorrelations ParentItem)
            {
                StringBuilder header = new StringBuilder();
                StringBuilder values = new StringBuilder();

                header.Append(ParentItem.Periods.Count());
                header.Append(",");

                header.Append("PP");
                header.Append(",");

                header.Append(ParentItem.uID.ID);
                for (int i = 0; i < ParentItem.Periods.Count(); i++)
                {
                    header.Append(",");
                    header.Append(ParentItem.Periods[i].pID.ID);
                }
                for(int i = 0; i< ParentItem.Periods.Count(); i++)
                {
                    values.Append("&0,0");
                }
                
                return new CorrelationString_PP($"{header}{values}");
            }

            public static CorrelationString_DP ConstructFromParentItem_Duration(IHasDurationCorrelations ParentItem)
            {
                StringBuilder header = new StringBuilder();
                StringBuilder values = new StringBuilder();

                header.Append(ParentItem.SubEstimates.Count);
                header.Append(",");

                header.Append("DP");
                header.Append(",");

                header.Append(ParentItem.uID.ID);
                for (int i = 0; i < ParentItem.SubEstimates.Count; i++)
                {
                    header.Append(",");
                    header.Append(ParentItem.SubEstimates[i].uID.ID);
                }
                for (int i = 0; i < ParentItem.SubEstimates.Count; i++)
                {
                    values.Append("&0,0");
                }
                return new CorrelationString_DP($"{header}{values}");
            }

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
                    case "CP":
                        return CorrelStringType.CostPair;
                    case "PM":
                        return CorrelStringType.PhasingMatrix;
                    case "PP":
                        return CorrelStringType.PhasingPair;
                    case "DP":
                        return CorrelStringType.DurationPair;
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
                    case CorrelStringType.PhasingPair:
                        return ParsePhasingTriple(stringValues);
                    case CorrelStringType.PhasingMatrix:
                        return ParsePhasingMatrix(stringValues);
                    case CorrelStringType.CostPair:
                        return ParseInputsTriple(stringValues);
                    case CorrelStringType.CostMatrix:
                        return ParseInputsMatrix(stringValues);
                    case CorrelStringType.DurationPair:
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

            protected virtual bool Validate(string correlStringValue)
            {
                //act as a switch for sending a string to its proper subclass validation
                return true;
            }

            public static CorrelationString ConstructFromCorrelationSheet(Sheets.CorrelationSheet correlSheet)
            {
                //Need to pick up sub IDs
                string sheetTag = Convert.ToString(correlSheet.xlSheet.Cells[1, 1].value);
                switch(sheetTag)
                {
                    case "$CORRELATION_CP":
                        return new CorrelationString_CP((Sheets.CorrelationSheet_CP)correlSheet);
                    case "$CORRELATION_CM":
                        return new CorrelationString_CM((Sheets.CorrelationSheet_CM)correlSheet);
                    case "$CORRELATION_PP":
                        return new CorrelationString_PP((Sheets.CorrelationSheet_PP)correlSheet);
                    case "$CORRELATION_DP":
                        return new CorrelationString_DP((Sheets.CorrelationSheet_DP)correlSheet);
                    case "$CORRELATION_DM":
                        return new CorrelationString_DM((Sheets.CorrelationSheet_DM)correlSheet);
                    default:
                        throw new Exception("Malformed correlation string");
                }
            }

            public static CorrelationString ConstructFromStringValue(string correlStringValue)
            {
                string[] lines = DelimitString(correlStringValue);
                string[] header = lines[0].Split(',');
                if (header.Length < 2)
                    throw new FormatException("Malformed correlation string");
                else
                {
                    switch (header[1])
                    {
                        //Each of these need to do their own integrity testing of the string value before they hand back an object
                        case "CP":
                            return new CorrelationString_CP(correlStringValue);
                        case "CM":
                            return new CorrelationString_CM(correlStringValue);
                        case "PP":
                            return new CorrelationString_PP(correlStringValue);
                        case "DP":
                            return new CorrelationString_DP(correlStringValue);
                        case "DM":
                            return new CorrelationString_DM(correlStringValue);
                        default:
                            throw new FormatException("Malformed correlation string");
                    }
                }
            }

            public static CorrelationString ConstructDefaultFromCostSheet(IHasSubs item, CorrelStringType csType)        //Construct default correlation string for estimate
            {
                switch (csType)
                {
                    case CorrelStringType.PhasingPair:
                        string[] start_dates = ((IHasPhasingCorrelations)item).Periods.Select(x => x.Start_Date).ToArray();
                        PairSpecification pairs = PairSpecification.ConstructFromSinglePair(start_dates.Count(), 0, 0);
                        return new Data.CorrelationString_PP(pairs, start_dates, item.uID.ID);
                    case CorrelStringType.CostPair:
                        if (((IHasCostCorrelations)item).SubEstimates.Count < 2)
                            return null;
                        PairSpecification ps = PairSpecification.ConstructFromSinglePair(((IHasCostCorrelations)item).SubEstimates.Count(), 0, 0); //new Triple(item.uID.ID, "0,0,0");
                        IEnumerable<string> fields = from ISub sub in ((IHasCostCorrelations)item).SubEstimates select sub.Name;        //need to print names, but get them from IDs?
                        return new Data.CorrelationString_CP(fields.ToArray(), ps, item.uID.ID, ((IHasCostCorrelations)item).SubEstimates.Select(x => x.uID.ID).ToArray());
                    case CorrelStringType.CostMatrix:
                        if (((IHasCostCorrelations)item).SubEstimates.Count < 2)
                            return null;
                        IEnumerable<string> fields2 = from Estimate_Item sub in item.ContainingSheetObject.GetSubEstimates(item.xlRow) select sub.Name;
                        return CorrelationString_CM.ConstructZeroString(fields2.ToArray());
                    case CorrelStringType.DurationMatrix:
                        throw new NotImplementedException();
                    case CorrelStringType.DurationPair:
                        if (((IHasDurationCorrelations)item).SubEstimates.Count < 2)
                            return null;
                        
                        IEnumerable<string> fields3 = from ISub sub in ((IHasDurationCorrelations)item).SubEstimates select sub.Name;
                        PairSpecification it2 = PairSpecification.ConstructFromSinglePair(fields3.Count(), 0, 0);
                        return new Data.CorrelationString_DP(fields3.ToArray(), it2, item.uID.ID);

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

            public static string GetHeaderFromParentItem(IHasSubs parentItem, SheetType correlationType)
            {
                switch (correlationType)
                {
                    case SheetType.Correlation_CM:
                        return parentItem.SubEstimates.First().xlCorrelCell_Cost.Value;
                    case SheetType.Correlation_CP:
                        return parentItem.SubEstimates.First().xlCorrelCell_Cost.Value;
                    case SheetType.Correlation_DM:
                        return parentItem.SubEstimates.First().xlCorrelCell_Duration.Value;
                    case SheetType.Correlation_DP:
                        return parentItem.SubEstimates.First().xlCorrelCell_Duration.Value;
                    default:
                        throw new Exception("Invalid correlation type for this method");
                }                
            }

            public static int GetNumberOfInputsFromCorrelStringValue(object correlString)
            {
                string cs = correlString.ToString();
                string[] delimited = cs.Split(',');
                return Convert.ToInt32(delimited[0]);
            }

            public static string GetParentIdFromCorrelStringValue(object correlString)
            {
                string cs = correlString.ToString();
                string[] delimited = cs.Split(',');
                return delimited[2];
            }

            public static string GetTypeOfCorrelationFromCorrelStringValue(object correlString)
            {
                string cs = correlString.ToString();
                string[] delimited = cs.Split(',');
                return delimited[1];
            }

            public static string GetParentIDFromCorrelStringValue(object correlString)
            {
                string cs = correlString.ToString();
                string[] lines = cs.Split('&');
                string[] delimited = lines[0].Split(',');
                return delimited[2];
            }
            #endregion
            public void PrintHeader(Excel.Range xlHeaderCell)
            {
                xlHeaderCell.Value = this.GetHeader();
            }
        }
    }
    
}
