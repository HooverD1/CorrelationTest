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

            public EstimateSheet(Excel.Worksheet xlSheet) : base(xlSheet) { }

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
                    if (Items[index].xlTypeCell.Value == "CE")
                    {
                        int input_index = index;
                        IHasCostCorrelations parentItem = (IHasCostCorrelations)Items[input_index];
                        while (input_index < Items.Count - 1)
                        {
                            ISub thisItem = (ISub)Items[++input_index];
                            if (thisItem.xlTypeCell.Value != "I")           //This needs to pick up the subestimate level for joint estimates, not just inputs..
                            {
                                index = input_index-1;
                                break;
                            }
                            else
                            {
                                parentItem.SubEstimates.Add(thisItem);
                                if(thisItem.Parent == null)
                                    thisItem.Parent = parentItem;
                            }
                        }
                    }
                    else if (Items[index].xlTypeCell.Value == "CASE")
                    {
                        int input_index = index;
                        IHasCostCorrelations parentItem = (IHasCostCorrelations)Items[input_index];
                        if (Items[input_index + 1] is ScheduleEstimate)
                        {
                            ISub thisItem = (ISub)Items[++input_index];
                            parentItem.SubEstimates.Add(thisItem);
                            if (thisItem.Parent == null)
                                thisItem.Parent = parentItem;
                        }
                        else
                            throw new Exception("Malformed CASE estimate");
                        
                        while (input_index < Items.Count - 1)
                        {
                            ISub thisItem = (ISub)Items[++input_index];
                            if (thisItem.xlTypeCell.Value != "I")
                            {
                                break;
                            }
                            else
                            {
                                parentItem.SubEstimates.Add(thisItem);
                                if (thisItem.Parent == null)
                                    thisItem.Parent = parentItem;
                            }
                        }
                    }
                    else if (Items[index].xlTypeCell.Value == "SE")
                    {
                        int input_index = index;
                        IHasDurationCorrelations parentItem = (IHasDurationCorrelations)Items[input_index];
                        
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
                                if (thisItem.Parent == null)
                                    thisItem.Parent = parentItem;
                            }
                        }
                    }
                    else if(Items[index].xlTypeCell.Value == "SACE")
                    {
                        int input_index = index;
                        IHasDurationCorrelations parentItem = (IHasDurationCorrelations)Items[input_index];
                        //GRAB THE COST ESTIMATE
                        if (Items[input_index + 1] is CostEstimate)
                        {
                            ISub thisItem = (ISub)Items[++input_index];
                            parentItem.SubEstimates.Add(thisItem);
                            if (thisItem.Parent == null)
                                thisItem.Parent = parentItem;
                        }
                        else
                            throw new Exception("Malformed SACE estimate");
                        //GRAB THE REST OF THE INPUTS
                        while (true)
                        {
                            ISub thisItem = (ISub)Items[++input_index];
                            if (thisItem.xlTypeCell.Value != "I")
                            {

                                break;
                            }
                            else
                            {
                                parentItem.SubEstimates.Add(thisItem);
                                if (thisItem.Parent == null)
                                    thisItem.Parent = parentItem;
                            }
                        }
                    }
                }
            }
            
            public override void PrintDefaultCorrelStrings()
            {
                foreach (IHasSubs item in (from item in Items where item is IHasSubs select item))
                {
                    if(item is ISub && item is IHasSubs)
                    {
                        //If it's a sub, see if its parent is a joint estimate. If so skip printing correlation.
                        if(((ISub)item).Parent is IJointEstimate) { continue; }
                    }
                    if (item is IHasCostCorrelations)
                        ((IHasCostCorrelations)item).PrintCostCorrelString();
                    if (item is IHasPhasingCorrelations)
                        ((IHasPhasingCorrelations)item).PrintPhasingCorrelString();
                    if (item is IHasDurationCorrelations)
                        ((IHasDurationCorrelations)item).PrintDurationCorrelString();
                    //if (item is IJointEstimate)
                    //{
                    //    if (item is CostScheduleEstimate)
                    //        ((CostScheduleEstimate)item).scheduleEstimate.PrintDurationCorrelString();
                    //    else if (item is ScheduleCostEstimate)
                    //        ((ScheduleCostEstimate)item).costEstimate.PrintCostCorrelString();
                    //    else
                    //        throw new Exception("Unknown joint estimate type");
                    //}
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
