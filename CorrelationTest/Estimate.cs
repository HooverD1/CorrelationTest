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
        public double[] Dollars { get; set; }
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
        public Excel.Range xlCorrelCell { get; set; }
        public Excel.Range xlLevelCell { get; set; }
        public string WBS_String { get; set; }
        public Dictionary<Estimate, double> CorrelPairs { get; set; }      //store non-zero correlations by unique id

        public Estimate(Excel.Range itemRow, ICostSheet ContainingSheetObject = null)
        {
            this.dispCoords = DisplayCoords.ConstructDisplayCoords(Sheets.Sheet.GetSheetType(itemRow.Worksheet));
            
            //this.TemporalCorrelStringObj = new Data.CorrelationString_Inputs
            this.xlRow = itemRow;
            this.xlDollarCell = itemRow.Cells[1, dispCoords.Dollar_Offset];
            this.xlTypeCell = itemRow.Cells[1, dispCoords.Type_Offset];
            this.xlCorrelCell = itemRow.Cells[1, dispCoords.InputCorrel_Offset];
            this.xlNameCell = itemRow.Cells[1, dispCoords.Name_Offset];
            this.xlIDCell = itemRow.Cells[1, dispCoords.ID_Offset];
            this.xlDistributionCell = itemRow.Cells[1, dispCoords.Distribution_Offset];
            this.xlLevelCell = itemRow.Cells[1, dispCoords.Level_Offset];
            this.ContainingSheetObject = ContainingSheetObject;
            this.Dollars = LoadDollars();
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
            this.CorrelPairs = new Dictionary<Estimate, double>();
        }

        private double[] LoadDollars()
        {
            double[] dollars = new double[10];
            for(int d = 0; d < dollars.Length; d++)
            {
                dollars[d] = xlDollarCell.Offset[0, d].Value;
            }
            return dollars;
        }

        public bool Equals(Estimate estimate)       //check the ID to determine equality
        {
            return this.uID.Equals(estimate.uID) ? true : false;
        }

        public void LoadSubEstimates()      //Returns a list of sub-estimates for this estimate
        {
            Excel.Worksheet xlSheet = this.xlRow.Worksheet;
            List<Estimate> returnList = new List<Estimate>();

            Excel.Range firstCell = xlSheet.Cells[this.xlRow.Row, dispCoords.Type_Offset];
            Excel.Range lastCell = xlSheet.Cells[1000000, dispCoords.Type_Offset].End[Excel.XlDirection.xlUp];
            Excel.Range pullRange = xlSheet.Range[firstCell, lastCell];
            Excel.Range[] estRows = PullEstimates(pullRange);
            for (int next = 1; next < estRows.Count(); next++)
            {
                Estimate nextEstimate;
                Estimate existingEstimate = null;
                //search for sub-estimate
                nextEstimate = new Estimate(estRows[next].EntireRow);      //build temp sub-estimate
                //if(ContainingSheetObject != null)
                //    existingEstimate = (from Estimate est in Estimates where est.ID == nextEstimate.ID select est).First();
                if (existingEstimate != null)
                    nextEstimate = existingEstimate;        //If you find a matching ID, use that instead of keeping the temp
                if (nextEstimate.Level - 1 == this.Level)
                {
                    //sub-estimate
                    this.SubEstimates.Add(nextEstimate);
                    nextEstimate.ParentEstimate = this;
                }
                else if (nextEstimate.Level <= this.Level)
                {
                    LoadCorrelatedValues(this.ParentEstimate);
                    return;
                }
            }
            LoadCorrelatedValues(this.ParentEstimate);
        }
        private void LoadCorrelatedValues(Estimate parentEstimate)      //this only ran on expand before -- now runs on build
        {
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

        private Excel.Range[] PullEstimates(Excel.Range pullRange)
        {
            Excel.Worksheet xlSheet = pullRange.Worksheet;
            IEnumerable<Excel.Range> returnVal = from Excel.Range cell in pullRange.Cells
                                                 where Convert.ToString(cell.Value) == "E"
                                                 select cell;
            return returnVal.ToArray<Excel.Range>();
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

        private UniqueID CreateID()
        {
            return new UniqueID(this.xlRow.Worksheet.Name, this.Name);
        }

        public void PrintName()
        {
            this.xlNameCell.Value = this.Name;
        }


    }

}
