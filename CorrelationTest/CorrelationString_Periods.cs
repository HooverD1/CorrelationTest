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
        public class CorrelationString_Periods : CorrelationString
        {
            public CorrelationString_Periods(Excel.Range xlDollarRange, double defaultCorrel = 0)     //Create new -- default correlation
            {
                int numberOfPeriods = xlDollarRange.Columns.Count;
                string[] fields = new string[numberOfPeriods];
                for(int t=1; t<=numberOfPeriods; t++)
                {
                    fields[t - 1] = $"T{t}";
                }
                this.Value = ExtensionMethods.CleanStringLinebreaks(CreateValue_Zero(fields, defaultCorrel));
            }

            public CorrelationString_Periods(string correlString)
            {
                this.Value = ExtensionMethods.CleanStringLinebreaks(correlString);
            }

            public static CorrelationString_Periods CreateZeroString(string[] fields)
            {
                //Need to downcast csi 
                var csi = new CorrelationString(fields);
                return new CorrelationString_Periods(csi.Value);
            }

            public CorrelationString_Periods(Data.CorrelationMatrix matrix)
            {
                this.Value = ExtensionMethods.CleanStringLinebreaks(CreateValue(matrix.GetIDs(), matrix.GetMatrix())); 
            }

            public static Data.CorrelationString_Periods ConstructString(PeriodID[] ids, string sheet, Dictionary<Tuple<UniqueID, UniqueID>, double> correls = null)
            {
                Data.CorrelationString_Periods correlationString = (CorrelationString_Periods)CreateZeroString((from UniqueID id in ids select id.ID).ToArray());       //build zero string
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
                            if (correls.ContainsKey(new Tuple<UniqueID, UniqueID>(id2, id1)))
                            {
                                matrix.SetCorrelation(id2, id1, correls[new Tuple<UniqueID, UniqueID>(id2, id1)]);
                            }
                        }
                    }
                    //convert to a string
                    return new Data.CorrelationString_Periods(matrix);      //return modified zero matrix as correl string
                }
            }

            public override object[] GetFields()
            {
                string[] fields = this.DelimitString()[0].Split(',');
                return fields.ToArray<object>();
            }

            public override UniqueID GetParentID()
            {
                string[] lines = this.Value.Split('&');
                string[] header = lines[1].Split(',');
                return new UniqueID(header[3]);
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

            public override void Expand(Excel.Range xlSource)
            {
                var id = this.GetIDs()[0];
                //construct the correlSheet
                Sheets.CorrelationSheet correlSheet = new Sheets.CorrelationSheet(this, xlSource, new Data.CorrelSheetSpecs());
                //print the correlSheet                         //CorrelationSheet NEEDS NEW CONSTRUCTORS BUILT FOR NON-INPUTS
                correlSheet.PrintToSheet();
            }

            public override void PrintToSheet(Excel.Range xlCell)
            {
                xlCell.Value = this.Value;
                xlCell.NumberFormat = "\"Ph Correl\";;;\"PH_CORREL\"";
                xlCell.EntireColumn.ColumnWidth = 10;
            }

            protected override string CreateValue(UniqueID[] ids, object[,] correlArray)
            {
                correlArray = ExtensionMethods.ReIndexArray<object>(correlArray);
                StringBuilder sb = new StringBuilder();
                sb.Append($"{ids.Length},PM");
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
        }
    }
    
}
