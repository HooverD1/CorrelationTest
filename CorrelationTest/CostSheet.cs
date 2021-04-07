using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace CorrelationTest
{
    public enum CostItems
    {
        I,
        CASE,
        SACE,
        CE,
        SE,
        W,
        F,
        S,
        Null
    }

    public abstract class CostSheet : Sheet
    {
        protected DialogResult OverwriteRepeatedIDs { get; set; }
        public DisplayCoords Specs { get; set; }
        public List<Item> Items { get; set; }
        //protected SheetType sheetType{ get; set;}

        //EXPAND
        public CostSheet(Excel.Worksheet xlSheet)
        {
            SheetType sheetType = ExtensionMethods.GetSheetType(xlSheet);
            this.Specs = DisplayCoords.ConstructDisplayCoords(sheetType);
            this.xlSheet = xlSheet;
            LoadItems();

        }

        public virtual List<Item> GetItemRows() { throw new Exception("Failed override"); }
        public virtual void LinkItemRows() { throw new Exception("Failed override"); }
        public void LoadItems()
        {
            this.Items = GetItemRows();
            LinkItemRows();
            this.LoadCorrelStrings();       //This has to be done after linking so that it knows what the parent child relationships are
            this.LoadDistributions();
            //Create CorrelationStrings
        }
        public virtual List<ISub> GetSubEstimates(Excel.Range parentRow) { throw new Exception("Failed override"); }    //Is this junk?
        public virtual void PrintDefaultCorrelStrings() { throw new Exception("Failed override"); }

        public void LoadDistributions()
        {
            //Cycle the items and load Distributions based on the parent's interfaces
            foreach(Item item in this.Items)
            {
                if(item is Estimate_Item)
                {
                    string distString = SpecifiedDistribution.GetDistributionStringFromRange(((Estimate_Item)item).xlDistributionCell);
                    if (item is ISub)
                    {
                        if (((ISub)item) is IHasCostCorrelations)
                        {
                            ((IHasCostCorrelations)item).CostDistribution = new SpecifiedDistribution(distString);
                        }
                        if (((ISub)item) is IHasDurationCorrelations)    //Really I want to have a place to store the duration correlation, not a distribution?
                        {
                            ((IHasDurationCorrelations)item).DurationDistribution = new SpecifiedDistribution(distString);
                        }
                    }
                    if (item is IHasPhasingCorrelations)
                    {
                        ((IHasPhasingCorrelations)item).PhasingDistribution = new SpecifiedDistribution(distString);
                    }
                }               
            }
        }

        public void LoadCorrelStrings()
        {
            foreach (IHasSubs item in (from item in Items where item is IHasSubs select item))
            {
                if (item is IHasPhasingCorrelations)        //Not on inputs, but everything else?
                    ((IHasPhasingCorrelations)item).LoadPhasingCorrelString();
                if (item is ISub)
                {
                    if(((ISub)item).Parent != null)
                    {
                        //items that have subs, can be a sub, and have a recorded parent are joint estimate sub-estimates and do not have their own correlation string
                        continue;       
                    }
                }
                //If it's not a subestimate of a joint estimate, load cost and/or duration correlation
                if (item is IHasCostCorrelations)
                    ((IHasCostCorrelations)item).LoadCostCorrelString();
                if (item is IHasDurationCorrelations)
                    ((IHasDurationCorrelations)item).LoadDurationCorrelString();
            }
        }

        public virtual object[] Get_xlFields()
        {
            throw new NotImplementedException();
        }

        public virtual void BuildCorrelations()
        {
            throw new NotImplementedException();
        }

        protected virtual void PrintCorrel_Cost(IHasCostCorrelations estimate, Dictionary<Tuple<string, string>, double> inputTemp = null)
        {
            /*
             * This is being called when "Build" is run. 
             * 
             */
            if (estimate.SubEstimates.Count >= 2)
            {
                //DAVID
                //This has too many subestimates
                string[] subIDs = (from Estimate_Item est in estimate.SubEstimates select est.uID.ID).ToArray();
                //check if any of the subestimates have NonZeroCorrel entries

                //This is sending in too many IDs
                object[] fields = estimate.SubEstimates.Select(x => x.Name).ToArray();
                estimate.CostCorrelationString.PrintToSheet(estimate.xlCorrelCell_Cost);
            }
        }

        protected virtual void PrintCorrel_Phasing(IHasPhasingCorrelations estimate, Dictionary<Tuple<PeriodID, PeriodID>, double> inputTemp = null)
        {
            /*
             * The print methods on the sheet object are there to compile a list of estimates
             * The print methods on the estimates should handle printing out correl strings
             * 
             * This should take a list of all estimates, recently built, cycle them, and call their print method to print correl strings (List<Estimate>)
             * The saved values should already be loaded into the estimates
             */
            //PeriodID[] periodIDs = (from Period prd in estimate.Periods select prd.pID).ToArray();
            //Data.CorrelationString_PM CorrelationString_PM = Data.CorrelationString_PM.ConstructString(periodIDs, this.xlSheet.Name);
            Data.CorrelationString correlationString = Data.CorrelationString.ConstructFromStringValue(estimate.xlCorrelCell_Phasing.Value);
            estimate.PhasingCorrelationString.PrintToSheet(estimate.xlCorrelCell_Phasing);
        }

        public virtual Excel.Range[] PullEstimates(Excel.Range pullRange, CostItems costType) { throw new Exception("Failed override"); }
        public virtual Excel.Range[] PullEstimates(Excel.Range pullRange) { throw new Exception("Failed override"); }

        //EXPAND
        public static CostSheet ConstructFromXlCostSheet(Excel.Worksheet xlSheet)
        {
            CostSheet sheetObj;
            switch(ExtensionMethods.GetSheetType(xlSheet))
            {
                case SheetType.WBS:
                    sheetObj = new Sheets.WBSSheet(xlSheet);
                    break;
                case SheetType.Estimate:
                    sheetObj = new Sheets.EstimateSheet(xlSheet);
                    break;
                default:
                    throw new Exception("Not a cost sheet type.");
            }
            return sheetObj;
        }

        protected string[] GetFields(Excel.Range selection)
        {
            string selection_id = selection.EntireRow.Cells[1, this.Specs.ID_Offset].value;
            CorrelationType ctype = ExtensionMethods.GetCorrelationTypeFromLink(selection);
            Item selectedItem = (from Item item in this.Items where item.uID.ID == selection_id select item).First();
            return (from ISub sub in ((IHasSubs)selectedItem).SubEstimates select sub.Name).ToArray();
        }

        protected void CreateCorrelStrings()
        {
            //Take the linked items and build their correlation strings
        }

    }
}
