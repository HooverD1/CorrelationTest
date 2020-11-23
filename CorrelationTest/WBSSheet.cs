using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace CorrelationTest
{
    namespace Sheets
    {
        public class WBSSheet: Sheet, ICostSheet, IHas_xlFields
        {
            public List<IEstimate> Estimates { get; set; }

            public WBSSheet(Excel.Worksheet xlSheet)
            {
                this.xlSheet = xlSheet;
                LoadParentEstimates();      //loads this.estimates
            }
            private List<IEstimate> LoadEstimates()
            {
                List<IEstimate> returnList = new List<IEstimate>();
                int iLastCell = xlSheet.Range["A1000000"].End[Excel.XlDirection.xlUp].Row;
                Excel.Range[] estRows = PullEstimates($"B2:B{iLastCell}");
                for(int index = 0; index < estRows.Count(); index++)
                {
                    Estimate parentEstimate = new Estimate(estRows[index].EntireRow);                    
                    for (int next = index+1; next < estRows.Count(); next++)
                    {
                        Estimate nextEstimate = new Estimate(estRows[next].EntireRow);
                        if(nextEstimate.Level - 1 == parentEstimate.Level)
                        {
                            //sub-estimate
                            parentEstimate.SubEstimates.Add(nextEstimate);
                            nextEstimate.ParentEstimate = parentEstimate;
                        }
                        else if(nextEstimate.Level >= parentEstimate.Level)
                        {
                            break;
                        }
                    }
                    returnList.Add(parentEstimate);
                }
                return returnList;
            }
            public void LoadParentEstimates()
            {
                this.Estimates = LoadEstimates();             
            }            

            private Excel.Range[] PullEstimates(string typeRange)
            {
                Excel.Range typeColumn = xlSheet.Range[typeRange];
                IEnumerable<Excel.Range> returnVal =    from Excel.Range cell in typeColumn.Cells
                                                        where Convert.ToString(cell.Value) == "E"
                                                        select cell;
                return returnVal.ToArray<Excel.Range>();
            }
            public object[] Get_xlFields()
            {
                string[] Estimate_WBSNames = (from Estimate est in Estimates select est.Name).ToArray();
                return Array.ConvertAll<string, object>(Estimate_WBSNames, new Converter<string, object>(x => (object)x));
            }
            public Data.CorrelationString LoadCorrelation(string correlString)
            {
                throw new NotImplementedException();
            }
            public Sheets.CorrelationSheet LoadCorrelationSheet(Data.CorrelationString correlStringObj)
            {
                throw new NotImplementedException();
            }
            public void SetCorrelation(Data.CorrelationString correlStringObj)
            {
                throw new NotImplementedException();
            }
            public override void PrintToSheet()
            {
                throw new NotImplementedException();
            }
            public override bool Validate()
            {
                throw new NotImplementedException();
            }
            public void BuildDefaultCorrelations()
            {
                int maxDepth = (from Estimate est in this.Estimates select est.Level).Max();
                if (this.Estimates.Any())
                    Estimates[0].xlCorrelCell.EntireColumn.Clear();        //Clear the Correlation String column   -- how to save current values here?
                foreach (Estimate est in this.Estimates)
                    PrintCorrel(est);  //recursively build out children
            }
           
            private void PrintCorrel(Estimate estimate)
            {
                if (estimate.SubEstimates.Count >= 2)
                {
                    object[] subNames = (from Estimate est in estimate.SubEstimates select est.Name).ToArray<object>();
                    //check if any of the subestimates have NonZeroCorrel entries
                    Data.CorrelationString correlationString;
                    correlationString = Data.CorrelationString.ConstructString(subNames, estimate);
                    correlationString.PrintToSheet(estimate.xlCorrelCell);
                }
            }
            
        }
    }
}
