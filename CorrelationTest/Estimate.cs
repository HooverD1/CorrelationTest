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
        public ICostSheet ContainingSheetObject { get; set; }
        public Distribution EstimateDistribution { get; set; }
        public Dictionary<string, object> DistributionParameters { get; set; }
        public char Type { get; set; }
        public Estimate ParentEstimate { get; set; }
        public List<Estimate> SubEstimates { get; set; }
        public List<Estimate> Siblings { get; set; }
        public Excel.Range xlRow { get; set; }
        public Data.CorrelationString TemporalCorrelStringObj { get; set; }
        public Data.CorrelationString InputCorrelStringObj { get; set; }
        public string ID { get; set; }
        public int Level { get; set; }
        public string Name { get; set; }
        public Excel.Range xlCorrelCell { get; set; }
        public string WBS_String { get; set; }
        public Dictionary<Estimate, double> CorrelPairs { get; set; }      //store non-zero correlations by unique id

        public Estimate(Excel.Range itemRow, ICostSheet ContainingSheetObject = null)
        {
            this.ContainingSheetObject = ContainingSheetObject;
            this.DistributionParameters = new Dictionary<string, object>()
              { { "Type", itemRow.Cells[1, 5].Value },
                { "Param1", itemRow.Cells[1, 6].Value },
                { "Param2", itemRow.Cells[1, 7].Value },
                { "Param3", itemRow.Cells[1, 8].Value },
                { "Param4", itemRow.Cells[1, 9].Value },
                { "Param5", itemRow.Cells[1, 10].Value } };
            this.EstimateDistribution = new Distribution(this.DistributionParameters);
            this.SubEstimates = new List<Estimate>();
            this.xlRow = itemRow;
            this.Level = Convert.ToInt32(itemRow.Cells[1, 1].Value);
            this.Type = Convert.ToChar(itemRow.Cells[1, 2].Value);
            this.Name = Convert.ToString(itemRow.Cells[1, 3].Value);
            this.WBS_String = Convert.ToString(itemRow.Cells[1, 3].Value);
            this.xlCorrelCell = itemRow.Cells[1, 4];
            this.ID = GetID();
            this.CorrelPairs = new Dictionary<Estimate, double>();
        }

        public bool Equals(Estimate estimate)       //check the ID to determine equality
        {
            return this.ID == estimate.ID ? true : false;
        }

        public void LoadSubEstimates(Excel.Range parentRow)      //Returns a list of sub-estimates for this estimate
        {
            Excel.Worksheet xlSheet = parentRow.Worksheet;
            List<Estimate> returnList = new List<Estimate>();
            int iLastCell = xlSheet.Range["A1000000"].End[Excel.XlDirection.xlUp].Row;
            Excel.Range[] estRows = PullEstimates(xlSheet, $"B{parentRow.Row}:B{iLastCell}");
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
                else if (nextEstimate.Level >= this.Level)
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
                this.CorrelPairs.Add(sibling, parentMatrix.AccessArray(this.ID, sibling.ID));
            }
        }

        private Excel.Range[] PullEstimates(Excel.Worksheet xlSheet, string typeRange)
        {
            Excel.Range typeColumn = xlSheet.Range[typeRange];
            IEnumerable<Excel.Range> returnVal = from Excel.Range cell in typeColumn.Cells
                                                 where Convert.ToString(cell.Value) == "E"
                                                 select cell;
            return returnVal.ToArray<Excel.Range>();
        }

        public object[] GetSubEstimateFields()
        {
            object[] subFields = new object[this.SubEstimates.Count];
            int index = 0;
            foreach(Estimate est in this.SubEstimates)
            {
                subFields[index] = est.Name;
                index++;
            }
            return subFields;
        }

        private string GetID()
        {
            return $"{this.xlRow.Worksheet.Name}|{this.Name}";
        }



    }

}
