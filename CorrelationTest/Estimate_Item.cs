using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using Accord.Statistics.Distributions.Univariate;

namespace CorrelationTest
{
    public class Estimate_Item : Item, IHasInputSubs, IHasDurationSubs, IHasPhasingSubs, ISub
    {
        public DisplayCoords dispCoords { get; set; }
        public int PeriodCount { get; set; }
        public Period[] Periods { get; set; }
        public UniqueID uID { get; set; }
        public Distribution ItemDistribution { get; set; }
        public Dictionary<string, object> DistributionParameters { get; set; }
        public string Type { get; set; }
        public Estimate_Item ParentEstimate { get; set; }
        public List<ISub> SubEstimates { get; set; }
        public List<Estimate_Item> Siblings { get; set; }
        public Data.CorrelationString TemporalCorrelStringObj { get; set; }
        public Data.CorrelationString_IM InputCorrelStringObj { get; set; }
        public int Level { get; set; }
        public Excel.Range xlDollarCell { get; set; }
        public Excel.Range xlIDCell { get; set; }
        public Excel.Range xlDistributionCell { get; set; }
        public Excel.Range xlLevelCell { get; set; }
        public string WBS_String { get; set; }
        public Dictionary<Estimate_Item, double> CorrelPairs { get; set; }      //store non-zero correlations by unique id

        public Estimate_Item(Excel.Range itemRow, CostSheet ContainingSheetObject) : base(itemRow, ContainingSheetObject)
        {
            this.dispCoords = DisplayCoords.ConstructDisplayCoords(ExtensionMethods.GetSheetType(itemRow.Worksheet));
            this.PeriodCount = 5;
            this.xlRow = itemRow;
            this.xlDollarCell = itemRow.Cells[1, dispCoords.Dollar_Offset];
            this.xlTypeCell = itemRow.Cells[1, dispCoords.Type_Offset];
            this.xlCorrelCell_Inputs = itemRow.Cells[1, dispCoords.InputCorrel_Offset];
            this.xlCorrelCell_Periods = itemRow.Cells[1, dispCoords.PhasingCorrel_Offset];
            this.xlNameCell = itemRow.Cells[1, dispCoords.Name_Offset];
            this.xlIDCell = itemRow.Cells[1, dispCoords.ID_Offset];
            this.xlDistributionCell = itemRow.Cells[1, dispCoords.Distribution_Offset];
            this.xlLevelCell = itemRow.Cells[1, dispCoords.Level_Offset];
            this.ContainingSheetObject = ContainingSheetObject;
            this.DistributionParameters = new Dictionary<string, object>()
              { { "Type", xlDistributionCell.Offset[0,0].Value },
                { "Param1", xlDistributionCell.Offset[0,1].Value },
                { "Param2", xlDistributionCell.Offset[0,2].Value },
                { "Param3", xlDistributionCell.Offset[0,3].Value },
                { "Param4", xlDistributionCell.Offset[0,4].Value },
                { "Param5", xlDistributionCell.Offset[0,5].Value } };
            this.ItemDistribution = new Distribution(this.DistributionParameters);
            this.SubEstimates = new List<ISub>();

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
            LoadPeriods();
            this.CorrelPairs = new Dictionary<Estimate_Item, double>();
        }

        public void LoadSubEstimates()
        {
            this.SubEstimates = GetSubs();
        }

        private List<ISub> GetSubs()
        {
            List<ISub> subEstimates = new List<ISub>();
            //Get the number of inputs
            int inputCount = Convert.ToInt32(xlRow.Cells[1, ContainingSheetObject.Specs.Level_Offset].value);    //Get the number of inputs
            for (int i = 1; i <= inputCount; i++)
            {
                subEstimates.Add(new Estimate_Item(xlRow.Offset[i, 0].EntireRow, ContainingSheetObject));
            }
            return subEstimates;
        }

        public void LoadPeriods()
        {
            this.Periods = GetPeriods();
        }
        private Period[] GetPeriods()
        {
            double[] dollars = LoadDollars();
            Period[] periods = new Period[PeriodCount];
            for (int i = 0; i < periods.Length; i++)
            {
                periods[i] = new Period(this.uID, i + 1, dollars[i]);
            }
            return periods;
        }
        private double[] LoadDollars()
        {
            double[] dollars = new double[PeriodCount];
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
            Estimate_Item parentEstimate = this.ParentEstimate;
            if (parentEstimate == null) { return; }
            if (parentEstimate.ParentEstimate == null) { return; }
            Data.CorrelationMatrix parentMatrix = Data.CorrelationMatrix.ConstructNew(parentEstimate.ParentEstimate.InputCorrelStringObj);     //How to build the matrix?
            foreach (Estimate_Item sibling in ParentEstimate.SubEstimates)
            {
                if (sibling == this)
                    continue;
                this.Siblings.Add(sibling);
                //create the string >> create the matrix >> retrieve values & store
                this.CorrelPairs.Add(sibling, parentMatrix.AccessArray(this.uID, sibling.uID));
            }
        }

        public void LoadExistingCorrelations()      //useful?
        {
            if (this.xlCorrelCell_Inputs != null)
            {
                Data.CorrelationString_IM csi = new Data.CorrelationString_IM(xlCorrelCell_Inputs.Value);
                this.InputCorrelStringObj = csi;

            }
            if (this.xlCorrelCell_Periods != null)
            {
                Data.CorrelationString_PM csp = new Data.CorrelationString_PM(xlCorrelCell_Periods.Value);
                this.TemporalCorrelStringObj = csp;
            }
        }

        public UniqueID[] GetSubEstimateIDs()
        {
            UniqueID[] subIDs = new UniqueID[this.SubEstimates.Count];
            int index = 0;
            foreach (Estimate_Item est in this.SubEstimates)
            {
                subIDs[index] = est.uID;
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

        public void PrintInputCorrelString()
        {
            Data.CorrelationString inString = Data.CorrelationString.Construct(this, Data.CorrelStringType.InputsTriple);
            if (inString != null)
                inString.PrintToSheet(xlCorrelCell_Inputs);
        }
        public void PrintPhasingCorrelString()
        {
            Data.CorrelationString phString = Data.CorrelationString.Construct(this, Data.CorrelStringType.PhasingTriple);
            if (phString != null)
                phString.PrintToSheet(xlCorrelCell_Periods);
        }
        public void PrintDurationCorrelString()
        {

        }
    }

}
