using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using Accord.Statistics.Distributions.Univariate;

namespace CorrelationTest
{
    public class Estimate : IEstimate
    {
        private DisplayCoords dispCoords { get; set; }
        public Period[] Periods { get; set; }
        public UniqueID uID { get; set; }
        public ICostSheet ContainingSheetObject { get; set; }
        public Distribution EstimateDistribution { get; set; }
        public Dictionary<string, object> DistributionParameters { get; set; }
        public char Type { get; set; }
        public Estimate ParentEstimate { get; set; }
        public List<Estimate> SubEstimates { get; set; }
        public List<Estimate> Siblings { get; set; }
        public Data.CorrelationString TemporalCorrelStringObj { get; set; }
        public Data.CorrelationString_Inputs InputCorrelStringObj { get; set; }
        public int Level { get; set; }
        public string Name { get; set; }
        public Excel.Range xlRow { get; set; }
        public Excel.Range xlDollarCell { get; set; }
        public Excel.Range xlIDCell { get; set; }
        public Excel.Range xlTypeCell { get; set; }
        public Excel.Range xlNameCell { get; set; }
        public Excel.Range xlDistributionCell { get; set; }
        public Excel.Range xlCorrelCell_Inputs { get; set; }
        public Excel.Range xlCorrelCell_Periods { get; set; }
        public Excel.Range xlLevelCell { get; set; }
        public string WBS_String { get; set; }
        public Dictionary<Estimate, double> CorrelPairs { get; set; }      //store non-zero correlations by unique id

        public Estimate(Excel.Range itemRow, ICostSheet ContainingSheetObject)
        {
            this.dispCoords = DisplayCoords.ConstructDisplayCoords(ExtensionMethods.GetSheetType(itemRow.Worksheet));
            
            //this.TemporalCorrelStringObj = new Data.CorrelationString_Inputs
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
            this.EstimateDistribution = new Distribution(this.DistributionParameters);
            this.SubEstimates = new List<Estimate>();
            
            this.Level = Convert.ToInt32(xlLevelCell.Value);
            this.Type = Convert.ToChar(xlTypeCell.Value);
            this.Name = Convert.ToString(xlNameCell.Value);

            if (xlIDCell.Value == null)
            {
                this.uID = CreateID();
                this.uID.PrintToCell(xlIDCell);
            }
            else
                this.uID = new UniqueID(xlIDCell.Value);
            this.Periods = LoadPeriods();
            this.CorrelPairs = new Dictionary<Estimate, double>();
        }
        private Period[] LoadPeriods()
        {
            double[] dollars = LoadDollars();
            Period[] periods = new Period[10];
            for(int i = 0; i < periods.Length; i++)
            {
                periods[i] = new Period(this.uID, i + 1, dollars[i]);
            }
            return periods;
        }
        private double[] LoadDollars()
        {
            double[] dollars = new double[10];
            for(int d = 0; d < dollars.Length; d++)
            {
                dollars[d] = xlDollarCell.Offset[0, d].Value ?? 0;
            }
            return dollars;
        }

        public bool Equals(Estimate estimate)       //check the ID to determine equality
        {
            return this.uID.Equals(estimate.uID) ? true : false;
        }

        public void LoadSubEstimates()
        {
            this.SubEstimates = GetSubEstimates();
        }

        public List<Estimate> GetSubEstimates()     //Attach this to the sheet? Check sheet type?
        {
            Excel.Worksheet xlSheet = this.xlRow.Worksheet;
            SheetType sheetType = ExtensionMethods.GetSheetType(xlSheet);
            CostItem ci;
            switch (sheetType)
            {
                case SheetType.Estimate:
                    ci = CostItem.I;
                    break;
                case SheetType.WBS:
                    ci = CostItem.E;
                    break;
                default:
                    throw new Exception("Unexpected sheet type");
            }
            List<Estimate> subestimates = new List<Estimate>();

            Excel.Range firstCell = xlSheet.Cells[this.xlRow.Row+1, dispCoords.Type_Offset];
            //iterate until you find <= level
            Excel.Range lastCell = firstCell.Offset[1,0];
            int offset = 0;
            while(true)
            {
                offset++;
                if (firstCell.Offset[offset, 0].Value != ci.ToString())
                    break;
                else
                    lastCell = firstCell.Offset[offset, 0];
            }
            Excel.Range pullRange = xlSheet.Range[firstCell, lastCell];
            Excel.Range[] estRows = this.ContainingSheetObject.PullEstimates(pullRange, ci);
            for (int next = 0; next < estRows.Count(); next++)
            {
                //search for sub-estimate
                Estimate nextEstimate = new Estimate(estRows[next].EntireRow, this.ContainingSheetObject);      //build temp sub-estimate
                if (nextEstimate.Level - 1 == this.Level)
                {
                    //sub-estimate
                    subestimates.Add(nextEstimate);
                    nextEstimate.ParentEstimate = this;
                }
                else if (nextEstimate.Level <= this.Level)
                {
                    LoadCorrelatedValues();
                    return subestimates;
                }
            }
            LoadCorrelatedValues();
            return subestimates;
        }
        
        private void LoadCorrelatedValues()      //this only ran on expand before -- now runs on build
        {
            this.Siblings = new List<Estimate>();
            Estimate parentEstimate = this.ParentEstimate;
            if (parentEstimate == null) { return; }
            if (parentEstimate.ParentEstimate == null) { return; }
            Data.CorrelationMatrix parentMatrix = new Data.CorrelationMatrix(parentEstimate.ParentEstimate.InputCorrelStringObj);     //How to build the matrix?
            foreach (Estimate sibling in ParentEstimate.SubEstimates)
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
                Data.CorrelationString_Inputs csi = new Data.CorrelationString_Inputs(xlCorrelCell_Inputs.Value);
                this.InputCorrelStringObj = csi;
                
            }
            if (this.xlCorrelCell_Periods != null)
            {
                Data.CorrelationString_Periods csp = new Data.CorrelationString_Periods(xlCorrelCell_Periods.Value);
                this.TemporalCorrelStringObj = csp;
            }
        }

        public UniqueID[] GetSubEstimateIDs()
        {
            UniqueID[] subIDs = new UniqueID[this.SubEstimates.Count];
            int index = 0;
            foreach(Estimate est in this.SubEstimates)
            {
                subIDs[index] = est.uID;
                index++;
            }
            return subIDs;
        }

        public UniqueID CreateID()
        {
            return new UniqueID(this.xlRow.Worksheet.Name, this.Name);
        }

        public void PrintName()
        {
            this.xlNameCell.Value = this.Name;
        }


    }

}
