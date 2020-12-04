using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;

namespace CorrelationTest
{
    namespace Sheets
    {
        public class WBSSheet: Sheet, ICostSheet, IHas_xlFields
        {
            private DisplayCoords dc = DisplayCoords.ConstructDisplayCoords(SheetType.WBS);
            private const int LevelColumn = 2;
            public List<IEstimate> Estimates { get; set; }
            private DialogResult OverwriteRepeatedIDs { get; set; }     

            public WBSSheet(Excel.Worksheet xlSheet)
            {
                this.xlSheet = xlSheet;
                this.Estimates = LoadEstimates();
            }

            public List<IEstimate> LoadEstimates()
            {
                List<IEstimate> returnList = new List<IEstimate>();
                Excel.Range lastCell = xlSheet.Cells[1000000, dc.Type_Offset].End[Excel.XlDirection.xlUp];
                Excel.Range firstCell = xlSheet.Cells[2, dc.Type_Offset];
                Excel.Range pullRange = xlSheet.Range[firstCell, lastCell];
                Excel.Range[] estRows = PullEstimates(pullRange.Address);
                int maxDepth = Convert.ToInt32((from Excel.Range row in estRows select row.Cells[1, LevelColumn].value).Max());

                for (int i = 1; i <= maxDepth; i++)
                {
                    Excel.Range[] topLevels = (from Excel.Range row in estRows where row.Cells[1, LevelColumn].value == i select row).ToArray<Excel.Range>();
                    for (int index = 0; index < topLevels.Count(); index++)
                    {
                        Estimate parentEstimate = new Estimate(topLevels[index].EntireRow);
                        parentEstimate.LoadSubEstimates();
                        returnList.Add(parentEstimate);
                    }
                }
                return returnList;
            }            
            
            private Excel.Range[] PullEstimates(string typeRange)       //return an array of rows
            {
                Excel.Range typeColumn = xlSheet.Range[typeRange];
                IEnumerable<Excel.Range> returnVal =    from Excel.Range cell in typeColumn.Cells
                                                        where Convert.ToString(cell.Value) == "E"
                                                        select cell.EntireRow;
                return returnVal.ToArray<Excel.Range>();
            }
            public object[] Get_xlFields()
            {
                string[] Estimate_WBSNames = (from Estimate est in Estimates select est.Name).ToArray();
                return Array.ConvertAll<string, object>(Estimate_WBSNames, new Converter<string, object>(x => (object)x));
            }
            public Data.CorrelationString LoadCorrelation(string correlString)
            {
                throw new NotImplementedException();
            }
            public Sheets.CorrelationSheet LoadCorrelationSheet(Data.CorrelationString correlStringObj)
            {
                throw new NotImplementedException();
            }
            public void SetCorrelation(Data.CorrelationString correlStringObj)
            {
                throw new NotImplementedException();
            }
            public override void PrintToSheet()
            {
                throw new NotImplementedException();
            }
            public override bool Validate()
            {
                throw new NotImplementedException();
            }
            public void BuildCorrelations()
            {
                //Input correlation
                int maxDepth = (from Estimate est in this.Estimates select est.Level).Max();
                var correlTemp = BuildCorrelTemp(this.Estimates);
                if(Estimates.Any())
                    Estimates[0].xlCorrelCell.EntireColumn.Clear();
                foreach (Estimate est in this.Estimates)
                {
                    PrintCorrel(est, correlTemp);  //recursively build out children
                }

                //Period correlation
                foreach(Estimate est in this.Estimates)
                {

                }

            }

            private Dictionary<Tuple<UniqueID, UniqueID>, double> BuildCorrelTemp(List<IEstimate> Estimates)
            {
                var correlTemp = new Dictionary<Tuple<UniqueID, UniqueID>, double>();   //<ID, ID>, correl_value
                if (this.Estimates.Any())
                {
                    //Save off existing correlations
                    //Create a correl string from the column
                    foreach (Estimate estimate in this.Estimates)
                    {
                        if (estimate.SubEstimates.Count == 0)
                            continue;
                        Data.CorrelationString_Inputs correlString;
                        if (estimate.xlCorrelCell.Value == null)        //No correlation string exists
                            correlString = Data.CorrelationString_Inputs.ConstructString(estimate.GetSubEstimateIDs(), this.xlSheet.Name);     //construct zero string
                        else
                            correlString = new Data.CorrelationString_Inputs(estimate.xlCorrelCell.Value);       //construct from string
                        var correlMatrix = new Data.CorrelationMatrix(correlString);
                        var matrixIDs = correlMatrix.GetIDs();
                        foreach (UniqueID id1 in matrixIDs)
                        {
                            foreach (UniqueID id2 in matrixIDs)
                            {
                                var newKey = new Tuple<UniqueID, UniqueID>(id1, id2);
                                if (!correlTemp.ContainsKey(newKey))
                                    correlTemp.Add(newKey, correlMatrix.AccessArray(id1, id2));
                            }
                        }
                    }
                    if (OverwriteRepeatedIDs == DialogResult.Yes)       //rebuild correlations
                        this.BuildCorrelations();
                }
                return correlTemp;
            }

            private void PrintCorrel(Estimate estimate, Dictionary<Tuple<UniqueID, UniqueID>, double> inputTemp = null)
            {
                if (estimate.SubEstimates.Count >= 2)
                {
                    
                    UniqueID[] subIDs = (from Estimate est in estimate.SubEstimates select est.uID).ToArray<UniqueID>();
                    //check if any of the subestimates have NonZeroCorrel entries
                    Data.CorrelationString_Inputs correlationString_inputs = Data.CorrelationString_Inputs.ConstructString(subIDs, this.xlSheet.Name, inputTemp);
                    correlationString_inputs.PrintToSheet(estimate.xlCorrelCell);

                    PeriodID[] periodIDs = (from Period prd in estimate.Periods select prd.pID).ToArray();
                    Data.CorrelationString_Periods correlationString_periods = Data.CorrelationString_Periods.ConstructString(periodIDs, this.xlSheet.Name, inputTemp);
                    //Data.CorrelationString_Periods correlationString_periods = Data.CorrelationString_Inputs.
                    //Need to be able to construct a string for _Periods
                }
            }
            
        }
    }
}
