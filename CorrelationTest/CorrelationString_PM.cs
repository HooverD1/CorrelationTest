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
        public class CorrelationString_PM : CorrelationString
        {
            public CorrelationString_PM(Excel.Range xlDollarRange, double defaultCorrel = 0)     //Create new -- default correlation
            {
                int numberOfPeriods = xlDollarRange.Columns.Count;
                string[] fields = new string[numberOfPeriods];
                for(int t=1; t<=numberOfPeriods; t++)
                {
                    fields[t - 1] = $"T{t}";
                }
                this.Value = ExtensionMethods.CleanStringLinebreaks(CreateValue_Zero(fields, defaultCorrel));
            }

            public CorrelationString_PM(string correlString)
            {
                this.Value = ExtensionMethods.CleanStringLinebreaks(correlString);
            }

            public CorrelationString_PM(Sheets.CorrelationSheet_Phasing correlSheet)
            {

            }


            private CorrelationString_PM(string[] start_dates)     //Zero string constructor
            {
                StringBuilder sb = new StringBuilder();
                sb.Append(start_dates.Length);
                sb.Append(",");
                sb.Append("PM");
                //No parent uid for a PM string
                sb.AppendLine();
                //Append the start dates
                for (int i = 0; i < start_dates.Length-1; i++)
                {
                    sb.Append(start_dates[i]);
                    sb.Append(",");
                }
                sb.Append(start_dates[start_dates.Length - 1]);
                sb.AppendLine();
                //Append default values (zeroes)
                for(int row = 0; row < start_dates.Length - 1; row++)
                {
                    for (int i = row; i < start_dates.Length - 2; i++)
                    {
                        sb.Append("0,");
                    }
                    sb.Append("0");
                    sb.AppendLine();
                }
                this.Value = sb.ToString();
            }

            public static CorrelationString_PM ConstructZeroString(string[] start_dates)
            {
                return new CorrelationString_PM(start_dates);
            }

            public CorrelationString_PM(Data.CorrelationMatrix matrix, string parentID)
            {
                this.Value = ExtensionMethods.CleanStringLinebreaks(CreateValue(parentID, matrix.GetFields(), matrix.GetMatrix())); 
            }

            public static bool Validate()
            {
                return true;
            }

            public static Data.CorrelationString_PM ConstructString(string parentID, string[] ids, string sheet, Dictionary<Tuple<string, string>, double> correls = null)
            {
                Data.CorrelationString_PM correlationString = (CorrelationString_PM)ConstructZeroString((from UniqueID id in ids select id.ID).ToArray());       //build zero string
                if (correls == null)
                    return correlationString;       //return zero string
                else
                {
                    Data.CorrelationMatrix matrix = Data.CorrelationMatrix.ConstructNew(correlationString);      //convert to zero matrix for modification
                    foreach (string id1 in ids)
                    {
                        foreach (string id2 in ids)
                        {
                            if (correls.ContainsKey(new Tuple<string, string>(id1, id2)))
                            {
                                matrix.SetCorrelation(id1, id2, correls[new Tuple<string, string>(id1, id2)]);
                            }
                            if (correls.ContainsKey(new Tuple<string, string>(id2, id1)))
                            {
                                matrix.SetCorrelation(id2, id1, correls[new Tuple<string, string>(id2, id1)]);
                            }
                        }
                    }
                    //convert to a string
                    return new Data.CorrelationString_PM(matrix, parentID);      //return modified zero matrix as correl string
                }
            }

            public override string[] GetFields()
            {
                string[] fields = CorrelationString.DelimitString(this.Value)[1].Split(',');
                return fields;
            }

            public override UniqueID GetParentID()
            {
                string[] lines = this.Value.Split('&');
                string[] header = lines[0].Split(',');
                return UniqueID.ConstructFromExisting(header[2]);
            }

            public override string[] GetIDs()
            {
                var period_ids = PeriodID.GeneratePeriodIDs(this.GetParentID(), this.GetNumberOfSubs());
                return period_ids.Select(x => x.ID).ToArray();
            }

            public override void Expand(Excel.Range xlSource)
            {
                var id = this.GetParentID();
                //construct the correlSheet
                Sheets.CorrelationSheet correlSheet = Sheets.CorrelationSheet.Construct(this, xlSource, new Data.CorrelSheetSpecs(SheetType.Correlation_PM));
                //print the correlSheet                         //CorrelationSheet NEEDS NEW CONSTRUCTORS BUILT FOR NON-INPUTS
                correlSheet.PrintToSheet();
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
                for (int i = 0; i < min; i++)
                {
                    xlFragments[i].Value = lines[i];
                    xlFragments[i].NumberFormat = "\"Ph Correl\";;;\"PH_CORREL\"";
                }
                xlFragments[0].EntireColumn.ColumnWidth = 10;
            }

            protected override string CreateValue(string parentID, object[] fields, object[,] correlArray)
            {
                correlArray = ExtensionMethods.ReIndexArray<object>(correlArray);
                StringBuilder sb = new StringBuilder();
                sb.Append($"{fields.Length},PM,{parentID}");
                sb.AppendLine();
                //for (int field = 0; field < correlArray.GetLength(1); field++)
                //{
                //    //Add fields
                //    sb.Append(fields[field]);
                //    if (field < correlArray.GetLength(1) - 1)
                //        sb.Append(",");
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
        }
    }
    
}
