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
        public class WBSSheet: CostSheet
        {
            private const SheetType sheetType = SheetType.WBS;
            
            public WBSSheet(Excel.Worksheet xlSheet)
            {
                this.Specs = DisplayCoords.ConstructDisplayCoords(sheetType);
                this.xlSheet = xlSheet;
                LoadEstimates(false);
            }

            public override List<Estimate> GetEstimates(bool LoadSubs)      //Returns a list of estimate objects for estimates on the sheet... this should really link to estimates on an estimate sheet
            {
                List<Estimate> returnList = new List<Estimate>();
                Excel.Range lastCell = xlSheet.Cells[1000000, Specs.Type_Offset].End[Excel.XlDirection.xlUp];
                Excel.Range firstCell = xlSheet.Cells[2, Specs.Type_Offset];
                Excel.Range pullRange = xlSheet.Range[firstCell, lastCell];
                Excel.Range[] estRows = PullEstimates(pullRange.Address);
                int maxDepth = Convert.ToInt32((from Excel.Range row in estRows select row.Cells[1, LevelColumn].value).Max());

                for (int i = 1; i <= maxDepth; i++)
                {
                    Excel.Range[] topLevels = (from Excel.Range row in estRows where row.Cells[1, LevelColumn].value == i select row).ToArray<Excel.Range>();
                    for (int index = 0; index < topLevels.Count(); index++)
                    {
                        Estimate parentEstimate = new Estimate(topLevels[index].EntireRow, this);
                        if(LoadSubs)
                            parentEstimate.LoadSubEstimates();
                        returnList.Add(parentEstimate);
                    }
                }
                return returnList;
            }            
            

            public override object[] Get_xlFields()
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
            public override void BuildCorrelations()
            {
                BuildCorrelations_Input();
                BuildCorrelations_Periods();
            }

            private void BuildCorrelations_Input()
            {
                //Input correlation
                int maxDepth = (from Estimate est in this.Estimates select est.Level).Max();
                var correlTemp = BuildCorrelTemp(this.Estimates);
                if (Estimates.Any())
                    Estimates[0].xlCorrelCell_Inputs.EntireColumn.Clear();
                foreach (Estimate est in this.Estimates)
                {
                    PrintCorrel_Inputs(est, correlTemp);  //recursively build out children
                }
            }

            private void BuildCorrelations_Periods()
            {
                //Period correlation
                foreach (Estimate est in this.Estimates)
                {
                    //Save the existing values
                    if (est.xlCorrelCell_Periods != null)
                    {
                        est.xlCorrelCell_Periods.Clear();
                    }
                    
                    PrintCorrel_Periods(est);
                }
            }

            private Dictionary<Tuple<UniqueID, UniqueID>, double> BuildCorrelTemp(List<Estimate> Estimates)
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
                        if (estimate.xlCorrelCell_Inputs.Value == null)        //No correlation string exists
                            correlString = Data.CorrelationString_Inputs.ConstructString(estimate.GetSubEstimateIDs(), this.xlSheet.Name);     //construct zero string
                        else
                            correlString = new Data.CorrelationString_Inputs(estimate.xlCorrelCell_Inputs.Value);       //construct from string
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

            protected override void PrintCorrel_Inputs(Estimate estimate, Dictionary<Tuple<UniqueID, UniqueID>, double> inputTemp = null)
            {
                /*
                 * This is being called when "Build" is run. 
                 * 
                 */
                if (estimate.SubEstimates.Count >= 2)
                {
                    
                    UniqueID[] subIDs = (from Estimate est in estimate.SubEstimates select est.uID).ToArray<UniqueID>();
                    //check if any of the subestimates have NonZeroCorrel entries
                    Data.CorrelationString_Inputs correlationString_inputs = Data.CorrelationString_Inputs.ConstructString(subIDs, this.xlSheet.Name, inputTemp);
                    correlationString_inputs.PrintToSheet(estimate.xlCorrelCell_Inputs);
                }
            }

            protected override void PrintCorrel_Periods(Estimate estimate, Dictionary<Tuple<PeriodID, PeriodID>, double> inputTemp = null)
            {
                /*
                 * The print methods on the sheet object are there to compile a list of estimates
                 * The print methods on the estimates should handle printing out correl strings
                 * 
                 * This should take a list of all estimates, recently built, cycle them, and call their print method to print correl strings (List<Estimate>)
                 * The saved values should already be loaded into the estimates
                 */
                PeriodID[] periodIDs = (from Period prd in estimate.Periods select prd.pID).ToArray();
                //Data.CorrelationString_Periods correlationString_periods = Data.CorrelationString_Periods.ConstructString(periodIDs, this.xlSheet.Name, inputTemp);
            }

            public override Excel.Range[] PullEstimates(Excel.Range pullRange, CostItem costType)
            {
                Excel.Worksheet xlSheet = pullRange.Worksheet;
                IEnumerable<Excel.Range> returnVal = from Excel.Range cell in pullRange.Cells
                                                     where Convert.ToString(cell.Value) == costType.ToString()
                                                     select cell;
                return returnVal.ToArray<Excel.Range>();
            }
        }
    }
}
