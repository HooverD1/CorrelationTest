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
                //LoadEstimates(false);
            }

            public override List<Item> GetItemRows(bool LoadSubs)      //Returns a list of estimate objects for estimates on the sheet... this should really link to estimates on an estimate sheet
            {
                List<Item> returnList = new List<Item>();
                Excel.Range lastCell = xlSheet.Cells[1000000, Specs.Type_Offset].End[Excel.XlDirection.xlUp];
                Excel.Range firstCell = xlSheet.Cells[2, Specs.Type_Offset];
                Excel.Range pullRange = xlSheet.Range[firstCell, lastCell];
                Excel.Range[] estRows = PullEstimates(pullRange);
                IEnumerable<object> levelObjects = from Excel.Range row in estRows select row.Cells[1, Specs.Level_Offset].value;
                List<int> levels = new List<int>();
                foreach(object level in levelObjects)
                {
                    if (int.TryParse(level.ToString(), out int parsedInt))
                        levels.Add(parsedInt);
                }
                int maxDepth = levels.Max();

                for (int i = 1; i <= maxDepth; i++)
                {
                    Excel.Range[] topLevels = (from Excel.Range row in estRows where row.Cells[1, Specs.Level_Offset].value == i select row).ToArray<Excel.Range>();
                    for (int index = 0; index < topLevels.Count(); index++)
                    {
                        Item parentRow = Item.Construct(topLevels[index].EntireRow, this);
                        if (LoadSubs && parentRow is IHasInputSubs)
                            ((IHasInputSubs)parentRow).LoadSubEstimates();
                        returnList.Add(parentRow);
                    }
                }
                return returnList;
            }
            
            public override List<ISub> GetSubEstimates(Excel.Range parentRow)     //Attach this to the sheet? Check sheet type?
            {
                Excel.Worksheet xlSheet = parentRow.Worksheet;
                SheetType sheetType = ExtensionMethods.GetSheetType(xlSheet);
                CostItems ci = CostItems.CE;
                List<ISub> subestimates = new List<ISub>();

                Excel.Range firstCell = xlSheet.Cells[parentRow.Row + 1, Specs.Type_Offset];
                //iterate until you find <= level
                Excel.Range lastCell = firstCell.Offset[1, 0];
                int offset = 0;
                while (true)
                {
                    offset++;
                    if (firstCell.Offset[offset, 0].Value != ci.ToString())
                        break;
                    else
                        lastCell = firstCell.Offset[offset, 0];
                }
                Excel.Range pullRange = xlSheet.Range[firstCell, lastCell];
                Excel.Range[] estRows = PullEstimates(pullRange);
                for (int next = 0; next < estRows.Count(); next++)
                {
                    Estimate_Item nextEstimate = (Estimate_Item)Item.Construct(estRows[next].EntireRow, this);      //build temp sub-estimate
                    subestimates.Add(nextEstimate);
                }
                return subestimates;
            }

            public override object[] Get_xlFields()
            {
                string[] Estimate_WBSNames = (from Estimate_Item est in CostRows select est.Name).ToArray();
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
                int maxDepth = (from Estimate_Item est in this.CostRows select est.Level).Max();
                var correlTemp = BuildCorrelTemp();
                if (CostRows.Any())
                    CostRows[0].xlCorrelCell_Inputs.EntireColumn.Clear();
                foreach (Estimate_Item est in this.CostRows)
                {
                    PrintCorrel_Inputs(est, correlTemp);  //recursively build out children
                }
            }

            private void BuildCorrelations_Periods()
            {
                //Period correlation
                foreach (Estimate_Item est in this.CostRows)
                {
                    //Save the existing values
                    if (est.xlCorrelCell_Periods != null)
                    {
                        est.xlCorrelCell_Periods.Clear();
                    }
                    
                    PrintCorrel_Periods(est);
                }
            }

            private Dictionary<Tuple<UniqueID, UniqueID>, double> BuildCorrelTemp()
            {
                var correlTemp = new Dictionary<Tuple<UniqueID, UniqueID>, double>();   //<ID, ID>, correl_value
                if (this.CostRows.Any())
                {
                    //Save off existing correlations
                    //Create a correl string from the column
                    foreach (Estimate_Item estimate in this.CostRows)
                    {
                        if (estimate.SubEstimates.Count == 0)
                            continue;
                        Data.CorrelationString_IM correlString;
                        if (estimate.xlCorrelCell_Inputs.Value == null)        //No correlation string exists
                            correlString = Data.CorrelationString_IM.ConstructString(estimate.GetSubEstimateIDs(), this.xlSheet.Name);     //construct zero string
                        else
                            correlString = new Data.CorrelationString_IM(estimate.xlCorrelCell_Inputs.Value);       //construct from string
                        var correlMatrix = Data.CorrelationMatrix.ConstructNew(correlString);
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

            protected override void PrintCorrel_Inputs(Estimate_Item estimate, Dictionary<Tuple<UniqueID, UniqueID>, double> inputTemp = null)
            {
                /*
                 * This is being called when "Build" is run. 
                 * 
                 */
                if (estimate.SubEstimates.Count >= 2)
                {
                    
                    UniqueID[] subIDs = (from Estimate_Item est in estimate.SubEstimates select est.uID).ToArray<UniqueID>();
                    //check if any of the subestimates have NonZeroCorrel entries
                    Data.CorrelationString_IM CorrelationString_IM = Data.CorrelationString_IM.ConstructString(subIDs, this.xlSheet.Name, inputTemp);
                    CorrelationString_IM.PrintToSheet(estimate.xlCorrelCell_Inputs);
                }
            }

            protected override void PrintCorrel_Periods(Estimate_Item estimate, Dictionary<Tuple<PeriodID, PeriodID>, double> inputTemp = null)
            {
                /*
                 * The print methods on the sheet object are there to compile a list of estimates
                 * The print methods on the estimates should handle printing out correl strings
                 * 
                 * This should take a list of all estimates, recently built, cycle them, and call their print method to print correl strings (List<Estimate>)
                 * The saved values should already be loaded into the estimates
                 */
                PeriodID[] periodIDs = (from Period prd in estimate.Periods select prd.pID).ToArray();
                //Data.CorrelationString_PM CorrelationString_PM = Data.CorrelationString_PM.ConstructString(periodIDs, this.xlSheet.Name, inputTemp);
            }

            public override Excel.Range[] PullEstimates(Excel.Range pullRange)
            {
                //This shouldn't pull by costType.. 
                Excel.Worksheet xlSheet = pullRange.Worksheet;
                IEnumerable<Excel.Range> returnVal = from Excel.Range cell in pullRange.Cells
                                                     where Convert.ToString(cell.Value) == "E" || Convert.ToString(cell.Value) == "S"
                                                     select cell.EntireRow;
                return returnVal.ToArray<Excel.Range>();
            }
        }
    }
}
