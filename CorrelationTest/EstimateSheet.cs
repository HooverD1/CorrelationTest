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
                foreach (Estimate_Item est in this.Items)
                {
                    est.ContainingSheetObject.GetSubEstimates(est.xlRow);     //this is returning too many subestimates       DAVID
                    PrintCorrel_Inputs(est, correlTemp);  //recursively build out children
                }
            }

            private void BuildCorrelations_Periods()
            {
                //Period correlation
                foreach (Estimate_Item est in this.Items)
                {
                    //Save the existing values
                    //if (est.xlCorrelCell_Periods != null)
                    //{
                    //    est.xlCorrelCell_Periods.Clear();
                    //}

                    PrintCorrel_Periods(est);
                }
            }

            private Dictionary<Tuple<UniqueID, UniqueID>, double> BuildCorrelTemp(List<Item> Estimates)
            {
                var correlTemp = new Dictionary<Tuple<UniqueID, UniqueID>, double>();   //<ID, ID>, correl_value
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

            public override void LinkItemRows()
            {
                for (int index = 0; index < Items.Count; index++)
                {
                    int thisLevel = Items[index].Level;
                    int indexStart = index;
                    while (Items[indexStart++].Level < thisLevel)
                    {
                        if (Items[indexStart].Level == thisLevel - 1)
                        {
                            string sub_uid = Items[indexStart].uID.ID;
                            IEnumerable<Item> theseSubs = from Item item in Items where item.uID.ID == sub_uid select item;
                            if (theseSubs.Count() > 1)
                                throw new Exception("Duplicated ID");
                            else if (theseSubs.Any())
                                ((IHasInputSubs)Items[index]).SubEstimates.Add((ISub)Items[index]); //If it found it, it must be a sub
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
                    subEstimates.Add(new Estimate_Item(parentRow.Offset[i, 0].EntireRow, this));
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
                                                     where Convert.ToString(cell.Value) == "CE" || Convert.ToString(cell.Value) == "I"
                                                     select cell;
                return returnVal.ToArray<Excel.Range>();
            }
        }
    }

}
