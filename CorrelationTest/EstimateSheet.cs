using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace CorrelationTest
{
    namespace Sheets
    {
        public class EstimateSheet : CostSheet
        {
            private const SheetType sheetType = SheetType.Estimate;

            public EstimateSheet(Excel.Worksheet xlSheet)
            {
                this.Specs = DisplayCoords.ConstructDisplayCoords(sheetType);
                this.xlSheet = xlSheet;
                //LoadEstimates(false);
            }

            public override void BuildCorrelations()
            {
                BuildCorrelations_Input();
                BuildCorrelations_Periods();
            }

            private void BuildCorrelations_Input()
            {
                //Input correlation
                var correlTemp = BuildCorrelTemp(this.Estimates);
                if (Estimates.Any())
                    Estimates[0].xlCorrelCell_Inputs.EntireColumn.Clear();
                foreach (Estimate est in this.Estimates)
                {
                    est.ContainingSheetObject.GetSubEstimates(est.xlRow);     //this is returning too many subestimates       DAVID
                    PrintCorrel_Inputs(est, correlTemp);  //recursively build out children
                }
            }

            private void BuildCorrelations_Periods()
            {
                //Period correlation
                foreach (Estimate est in this.Estimates)
                {
                    //Save the existing values
                    //if (est.xlCorrelCell_Periods != null)
                    //{
                    //    est.xlCorrelCell_Periods.Clear();
                    //}

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
                        Data.CorrelationString_IM correlString;
                        if (estimate.xlCorrelCell_Inputs.Value == null)        //No correlation string exists
                            correlString = Data.CorrelationString_IM.ConstructString(estimate.GetSubEstimateIDs(), this.xlSheet.Name);     //construct zero string
                        else
                            correlString = new Data.CorrelationString_IM(estimate.xlCorrelCell_Inputs.Value);       //construct from string
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

            public override object[] Get_xlFields()
            {
                throw new NotImplementedException();
            }

            public override List<Estimate> GetEstimates(bool LoadSubs)
            {
                List<Estimate> returnList = new List<Estimate>();
                Excel.Range lastCell = xlSheet.Cells[1000000, Specs.Type_Offset].End[Excel.XlDirection.xlUp];
                Excel.Range firstCell = xlSheet.Cells[2, Specs.Type_Offset];
                Excel.Range pullRange = xlSheet.Range[firstCell, lastCell];
                Excel.Range[] estRows = PullEstimates(pullRange, CostItems.E);       //Pull the estimates (not the inputs)
                for (int index = 0; index < estRows.Count(); index++)
                {
                    Estimate parentEstimate = new Estimate(estRows[index].EntireRow, this);
                    if(LoadSubs)
                        parentEstimate.SubEstimates = this.GetSubEstimates(estRows[index].EntireRow);     //Get the subestimates for this parent row
                    returnList.Add(parentEstimate);
                }
                return returnList;
            }

            public override List<Estimate> GetSubEstimates(Excel.Range parentRow)
            {
                List<Estimate> subEstimates = new List<Estimate>();
                //Get the number of inputs
                int inputCount = Convert.ToInt32(parentRow.Cells[1, this.Specs.Level_Offset].value);    //Get the number of inputs
                for(int i = 1; i <= inputCount; i++)
                {
                    subEstimates.Add(new Estimate(parentRow.Offset[i, 0].EntireRow, this));
                }
                return subEstimates;
            }

            public override void PrintToSheet()
            {
                throw new NotImplementedException();
            }

            public override bool Validate()
            {
                throw new NotImplementedException();
            }

            public override Excel.Range[] PullEstimates(Excel.Range pullRange, CostItems costType)
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
