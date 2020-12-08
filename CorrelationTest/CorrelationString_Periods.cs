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
                this.Value = CreateValue_Zero(fields, defaultCorrel);
            }

            public CorrelationString_Periods(string correlString)
            {
                this.Value = correlString;
            }

            public static CorrelationString_Periods CreateZeroString(string[] fields)
            {
                //Need to downcast csi 
                var csi = new CorrelationString(fields);
                return new CorrelationString_Periods(csi.Value);
            }

            public CorrelationString_Periods(Data.CorrelationMatrix matrix)
            {
                this.Value = CreateValue(matrix.GetIDs(), matrix.GetMatrix());
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

        }
    }
    
}
