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

            public EstimateSheet(Excel.Worksheet xlSheet) : base(xlSheet)
            {
                sheetType = SheetType.Estimate;
            }

            public override void BuildCorrelations()
            {
                BuildCorrelations_Input();
                BuildCorrelations_Periods();
            }

            private void BuildCorrelations_Input()
            {
                //Input correlation
                var correlTemp = BuildCorrelTemp(this.Items);
                if (Items.Any())
                    Items[0].xlCorrelCell_Inputs.EntireColumn.Clear();
                foreach (IHasInputSubs est in this.Items)
                {
                    est.ContainingSheetObject.GetSubEstimates(est.xlRow); 
                    PrintCorrel_Inputs(est, correlTemp);  //recursively build out children
                }
            }

            private void BuildCorrelations_Periods()
            {
                //Period correlation
                foreach (IHasPhasingSubs est in this.Items)
                {
                    PrintCorrel_Periods(est);
                }
            }

            private Dictionary<Tuple<string, string>, double> BuildCorrelTemp(List<Item> Estimates)
            {
                var correlTemp = new Dictionary<Tuple<string, string>, double>();   //<ID, ID>, correl_value
                if (this.Items.Any())
                {
                    //Save off existing correlations
                    //Create a correl string from the column
                    foreach (Estimate_Item estimate in this.Items)
                    {
                        if (estimate.SubEstimates.Count == 0)
                            continue;
                        Data.CorrelationString_IM correlString;
                        if (estimate.xlCorrelCell_Inputs.Value == null)        //No correlation string exists
                            correlString = Data.CorrelationString_IM.ConstructString(estimate.uID.ID, estimate.GetSubEstimateIDs(), estimate.SubEstimates.Select(x=>x.Name).ToArray(), this.xlSheet.Name);     //construct zero string
                        else
                            correlString = new Data.CorrelationString_IM(estimate.xlCorrelCell_Inputs.Value);       //construct from string
                        var correlMatrix = Data.CorrelationMatrix.ConstructNew(correlString);
                        string[] ids = Estimates.Select(x => x.uID.ID).ToArray();
                        foreach (string id1 in ids)
                        {
                            foreach (string id2 in ids)
                            {
                                var newKey = new Tuple<string, string>(id1, id2);
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

            public override List<Item> GetItemRows()
            {
                List<Item> returnList = new List<Item>();
                Excel.Range lastCell = xlSheet.Cells[1000000, Specs.Type_Offset].End[Excel.XlDirection.xlUp];
                Excel.Range firstCell = xlSheet.Cells[2, Specs.Type_Offset];
                Excel.Range pullRange = xlSheet.Range[firstCell, lastCell];
                Excel.Range[] estRows = PullEstimates(pullRange);      
                for (int index = 0; index < estRows.Count(); index++)
                {
                    var currentItem = Item.ConstructFromRow(estRows[index].EntireRow, this);
                    returnList.Add(currentItem);
                }
                return returnList;
            }

            public override void LinkItemRows()     //There are no levels on an estimate sheet!
            {
                for (int index = 0; index < Items.Count; index++)
                {
                    if (Items[index].xlTypeCell.Value == "CE")
                    {
                        IHasInputSubs parentItem = (IHasInputSubs)Items[index];
                        int input_index = index;
                        while (input_index < Items.Count - 1)
                        {
                            ISub thisItem = (ISub)Items[++input_index];
                            if (thisItem.xlTypeCell.Value != "I")
                            {
                                index = input_index-1;
                                break;
                            }
                            else
                            {
                                parentItem.SubEstimates.Add(thisItem);
                            }
                        }
                    }
                    else if (Items[index].xlTypeCell.Value == "SE")
                    {
                        IHasDurationSubs parentItem = (IHasDurationSubs)Items[index];
                        int input_index = index;
                        while (true)
                        {
                            ISub thisItem = (ISub)Items[++input_index];
                            if (thisItem.xlTypeCell.Value != "I")
                            {
                                index = input_index-1;
                                break;
                            }
                            else
                            {
                                parentItem.SubEstimates.Add(thisItem);
                            }
                        }
                    }
                }
            }

            public override void PrintDefaultCorrelStrings()
            {
                foreach (IHasSubs item in Items)
                {
                    if (item is IHasInputSubs)
                        ((IHasInputSubs)item).PrintInputCorrelString();
                    if (item is IHasPhasingSubs)
                        ((IHasPhasingSubs)item).PrintPhasingCorrelString();
                    if (item is IHasDurationSubs)
                        ((IHasDurationSubs)item).PrintDurationCorrelString();
                }
            }

            public override List<ISub> GetSubEstimates(Excel.Range parentRow)
            {
                List<ISub> subEstimates = new List<ISub>();
                //Get the number of inputs
                int inputCount = Convert.ToInt32(parentRow.Cells[1, this.Specs.Level_Offset].value);    //Get the number of inputs
                for(int i = 1; i <= inputCount; i++)
                {
                    subEstimates.Add((ISub)Item.ConstructFromRow(parentRow.Offset[i, 0].EntireRow, this));
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

            public override Excel.Range[] PullEstimates(Excel.Range pullRange)
            {
                Excel.Worksheet xlSheet = pullRange.Worksheet;
                IEnumerable<Excel.Range> returnVal = from Excel.Range cell in pullRange.Cells
                                                     where Convert.ToString(cell.Value) == "CE" || 
                                                           Convert.ToString(cell.Value) == "SE" ||
                                                           Convert.ToString(cell.Value) == "CASE" ||
                                                           Convert.ToString(cell.Value) == "SACE" ||
                                                           Convert.ToString(cell.Value) == "I"
                                                     select cell;
                return returnVal.ToArray<Excel.Range>();
            }
        }
    }

}
