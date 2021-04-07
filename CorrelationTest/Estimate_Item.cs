using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using Accord.Statistics.Distributions.Univariate;

namespace CorrelationTest
{
    public class Estimate_Item : Item, IHasSubs
    {
        public DisplayCoords dispCoords { get; set; }
        public Period[] Periods { get; set; }
        public IEstimateDistribution PhasingDistribution { get; set; }
        public IEstimateDistribution CostDistribution { get; set; } //Cost or Schedule
        public IEstimateDistribution DurationDistribution { get; set; }
        public Data.CorrelationString CostCorrelationString { get; set; }
        public Data.CorrelationString DurationCorrelationString { get; set; }
        public Data.CorrelationString PhasingCorrelationString { get; set; }
        public Dictionary<string, object> ValueDistributionParameters { get; set; }
        public Dictionary<string, object> PhasingDistributionParameters { get; set; }
        public string Type { get; set; }
        public IHasSubs Parent { get; set; }
        public List<ISub> SubEstimates { get; set; } = new List<ISub>();
        public List<Estimate_Item> Siblings { get; set; }
        public Data.CorrelationString CorrelStringObj_Cost { get; set; }
        public Data.CorrelationString CorrelStringObj_Phasing { get; set; }
        public Data.CorrelationString CorrelStringObj_Duration { get; set; }
        public Excel.Range xlDollarCell { get; set; }
        public Excel.Range xlIDCell { get; set; }
        public Excel.Range xlDistributionCell { get; set; }
        public string WBS_String { get; set; }
        public Dictionary<Estimate_Item, double> CorrelPairs { get; set; }      //store non-zero correlations by unique id

        public Estimate_Item(Excel.Range itemRow, CostSheet ContainingSheetObject) : base(itemRow, ContainingSheetObject)
        {
            this.dispCoords = DisplayCoords.ConstructDisplayCoords(ExtensionMethods.GetSheetType(itemRow.Worksheet));
            this.xlRow = itemRow;
            this.xlDollarCell = itemRow.Cells[1, dispCoords.Dollar_Offset];
            this.xlTypeCell = itemRow.Cells[1, dispCoords.Type_Offset];
            this.xlCorrelCell_Cost = itemRow.Cells[1, dispCoords.CostCorrel_Offset];
            this.xlCorrelCell_Duration = itemRow.Cells[1, dispCoords.DurationCorrel_Offset];
            this.xlCorrelCell_Phasing = itemRow.Cells[1, dispCoords.PhasingCorrel_Offset];
            this.xlNameCell = itemRow.Cells[1, dispCoords.Name_Offset];
            this.xlIDCell = itemRow.Cells[1, dispCoords.ID_Offset];
            this.xlDistributionCell = itemRow.Cells[1, dispCoords.Distribution_Offset];
            this.xlLevelCell = itemRow.Cells[1, dispCoords.Level_Offset];
            this.ContainingSheetObject = ContainingSheetObject;


            this.PhasingDistributionParameters = new Dictionary<string, object>() {
                { "Type", "Normal" },
                { "Param1", 1 },
                { "Param2", 1 },
                { "Param3", 1 },
                { "Param4", 0 },
                { "Param5", 0 } };
            this.PhasingDistribution = new SpecifiedDistribution(PhasingDistributionParameters);    //Should this even be a Distribution object? More of a schedule.

            this.Level = Convert.ToInt32(xlLevelCell.Value);
            this.Type = Convert.ToString(xlTypeCell.Value);
            this.Name = Convert.ToString(xlNameCell.Value);

            if (xlIDCell.Value == null)
            {
                this.uID = CreateID();
                this.uID.PrintToCell(xlIDCell);
            }
            else
                this.uID = UniqueID.ConstructFromExisting(xlIDCell.Value);
            LoadPhasing(xlRow);
            this.CorrelPairs = new Dictionary<Estimate_Item, double>();
        }

        public void LoadSubEstimates()
        {
            this.SubEstimates = GetSubs();
        }

        public string[] GetFields()
        {
            IEnumerable<string> fields = from ISub sub in SubEstimates select sub.Name;
            return fields.ToArray();
        }

        private List<ISub> GetSubs()        //This only works for the estimate sheet because it tells you the number of subs and they're all contiguous
        {
            List<ISub> subEstimates = new List<ISub>();
            //Get the number of inputs
            int inputCount = Convert.ToInt32(xlRow.Cells[1, ContainingSheetObject.Specs.Level_Offset].value);    //Get the number of inputs
            for (int i = 1; i <= inputCount; i++)
            {
                subEstimates.Add((ISub)Item.ConstructFromRow(xlRow.Offset[i, 0].EntireRow, ContainingSheetObject));
            }
            return subEstimates;
        }

        public void LoadPhasing(Excel.Range xlRow)
        {
            this.Periods = GetPeriods();
        }
        private Period[] GetPeriods()
        {
            double[] dollars = LoadDollars();
            Period[] periods = new Period[5];
            for (int i = 1; i <= periods.Length; i++)
            {
                periods[i-1] = new Period(this.uID, $"P{i}", dollars[i-1]);     //Need to be able to pull the dates off the sheet here.
            }
            return periods;
        }
        private double[] LoadDollars()
        {
            double[] dollars = new double[5];
            for (int d = 0; d < dollars.Length; d++)
            {
                dollars[d] = xlDollarCell.Offset[0, d].Value ?? 0;
            }
            return dollars;
        }

        public bool Equals(Estimate_Item estimate)       //check the ID to determine equality
        {
            return this.uID.Equals(estimate.uID) ? true : false;
        }

      

        private void LoadCorrelatedValues()      //this only ran on expand before -- now runs on build
        {
            this.Siblings = new List<Estimate_Item>();
            if (this.Parent == null) { return; }
            if (!(Parent is ISub)) { return; }
            if (((ISub)Parent).Parent == null) { return; }

            IHasSubs grandparent = ((ISub)Parent).Parent;

            SheetType correlType;
            if (grandparent is IHasCostCorrelations) { correlType = ((IHasCostCorrelations)grandparent).CostCorrelationString.GetCorrelType(); }
            else if(grandparent is IHasDurationCorrelations) { correlType = ((IHasDurationCorrelations)grandparent).DurationCorrelationString.GetCorrelType(); }
            else { correlType = SheetType.Unknown; }
            Data.CorrelationMatrix parentMatrix = Data.CorrelationMatrix.ConstructFromParentItem(grandparent, correlType, null);     //How to build the matrix?
            foreach (Estimate_Item sibling in Parent.SubEstimates)
            {
                if (sibling == this)
                    continue;
                this.Siblings.Add(sibling);
                //create the string >> create the matrix >> retrieve values & store
                this.CorrelPairs.Add(sibling, parentMatrix.AccessArray(this.uID.ID, sibling.uID.ID));
            }

        }

        public string[] GetSubEstimateIDs()
        {
            string[] subIDs = new string[this.SubEstimates.Count];
            int index = 0;
            foreach (Estimate_Item est in this.SubEstimates)
            {
                subIDs[index] = est.uID.ID;
                index++;
            }
            return subIDs;
        }

        public void LoadUID()
        {
            this.uID = GetUID();
        }

        protected virtual UniqueID GetUID()
        {
            if (this.xlRow.Cells[1, ContainingSheetObject.Specs.ID_Offset].value != null)
            {
                string idString = Convert.ToString(this.xlRow.Cells[1, ContainingSheetObject.Specs.ID_Offset].value);
                return UniqueID.ConstructFromExisting(idString);
            }
            else
            {
                //Create new ID
                return UniqueID.ConstructNew("E");
            }
        }

        public UniqueID CreateID()
        {
            return UniqueID.ConstructNew("E");
        }

        public void PrintName()
        {
            this.xlNameCell.Value = this.Name;
        }

        public virtual void LoadCostCorrelString()
        {
            //this.CostCorrelationString = Data.CorrelationString.ConstructDefaultFromCostSheet(this, Data.CorrelStringType.CostPair);
            //return;

            //This needs to check if a string already exists.
            //It checks in the first child since Duration correlation is stored against the child row
            if (!this.SubEstimates.Any())
                return;
            var firstChild = this.SubEstimates.First();
            if (firstChild.xlCorrelCell_Cost.Value == null)
            {
                //This should check the correl type to split pairwise from matrix
                this.CostCorrelationString = Data.CorrelationString.ConstructDefaultFromCostSheet(this, Data.CorrelStringType.CostPair);
            }
            else
            {
                //Something in the cell that can either be resolved into a correl string or not
                try
                {
                    string costString = Data.CorrelationString.ConstructStringFromRange(from ISub sub in this.SubEstimates select sub.xlCorrelCell_Cost);
                    this.CostCorrelationString = Data.CorrelationString.ConstructFromStringValue(costString);
                }
                catch
                {
                    if (MyGlobals.DebugMode)
                        throw new Exception("Malformed correl string");
                }
            }
        }

        public virtual void PrintCostCorrelString()
        {
            
            IEnumerable<Excel.Range> xlFragments = from ISub sub in this.SubEstimates
                                                   select sub.xlCorrelCell_Cost;
            if (this.CostCorrelationString != null)
                this.CostCorrelationString.PrintToSheet(xlFragments.ToArray());

        }

        public virtual void LoadPhasingCorrelString()
        {
            //Loading default
            if(xlCorrelCell_Phasing.Value == null)
                this.PhasingCorrelationString = Data.CorrelationString.ConstructDefaultFromCostSheet(this, Data.CorrelStringType.PhasingPair);
            else
                this.PhasingCorrelationString = Data.CorrelationString.ConstructFromStringValue(xlCorrelCell_Phasing.Value);
        }

        public virtual void PrintPhasingCorrelString()
        {
            if (this.PhasingCorrelationString != null)
                this.PhasingCorrelationString.PrintToSheet(xlCorrelCell_Phasing);        //Phasing goes on the parent. Cost and Dura go on the children
        }

        public virtual void LoadDurationCorrelString()
        {
            //This needs to check if a string already exists.
            //It checks in the first child since Duration correlation is stored against the child row
            var firstChild = this.SubEstimates.First();
            if(firstChild.xlCorrelCell_Duration.Value == null)
            {
                this.DurationCorrelationString = Data.CorrelationString.ConstructDefaultFromCostSheet(this, Data.CorrelStringType.DurationPair);
            }
            else
            {
                //Something in the cell that can either be resolved into a correl string or not
                try
                {
                    string durationString = Data.CorrelationString.ConstructStringFromRange(from ISub sub in this.SubEstimates select sub.xlCorrelCell_Duration);
                    this.DurationCorrelationString = Data.CorrelationString.ConstructFromStringValue(durationString);
                }
                catch
                {
                    if(MyGlobals.DebugMode)
                        throw new Exception("Malformed correl string");
                }
            }            
        }

        public virtual void PrintDurationCorrelString()
        {
            IEnumerable<Excel.Range> xlFragments = from ISub sub in this.SubEstimates
                                                   select sub.xlCorrelCell_Duration;
            if (this.DurationCorrelationString != null)
                this.DurationCorrelationString.PrintToSheet(xlFragments.ToArray());
        }

        public Data.CorrelationString GetCorrelationString(Data.CorrelStringType cst)
        {       //Build the CorrelationString from the existing fragments on the sheet
            IEnumerable<Excel.Range> fragments = from ISub sub in this.SubEstimates select sub.xlCorrelCell_Cost;
            string csValue = Data.CorrelationString.ConstructStringFromRange(fragments);
            return Data.CorrelationString.ConstructFromStringValue(csValue);
            //Get the fragment ranges
            //Feed them to CorrelationString
            //Return the CorrelString string
            //Build the CorrelationString object
        }

        public void Expand(CorrelationType correlType)
        {
            switch (correlType)
            {
                case CorrelationType.Cost:
                    Expand_Cost();
                    break;
                case CorrelationType.Phasing:
                    Expand_Phasing();
                    break;
                case CorrelationType.Duration:
                    Expand_Duration();
                    break;
            }
        }

        private void Expand_Cost()
        {
            SheetType typeOfCost = this.CostCorrelationString.GetCorrelType();
            Sheets.CorrelationSheet correlSheet = Sheets.CorrelationSheet.ConstructFromParentItem(this, typeOfCost);
            correlSheet.PrintToSheet();
        }

        private void Expand_Phasing()
        {
            SheetType typeOfCost = this.PhasingCorrelationString.GetCorrelType();
            Sheets.CorrelationSheet correlSheet = Sheets.CorrelationSheet.ConstructFromParentItem(this, typeOfCost);
            correlSheet.PrintToSheet();
        }

        private void Expand_Duration()      //Inefficiency: I believe all the items are already loaded when creationg correlSheet - .PrintToSheet() reloads the cost sheet, which reloads the items.
        {
            SheetType typeOfCost = this.DurationCorrelationString.GetCorrelType();
            Sheets.CorrelationSheet correlSheet = Sheets.CorrelationSheet.ConstructFromParentItem(this, typeOfCost);
            correlSheet.PrintToSheet();
        }
    }

}
