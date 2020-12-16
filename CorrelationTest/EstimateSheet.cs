﻿using System;
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
            private const SheetType sheetType = SheetType.Estimate;

            public EstimateSheet(Excel.Worksheet xlSheet)
            {
                this.LevelColumn = 2;
                this.dc = DisplayCoords.ConstructDisplayCoords(sheetType);
                this.xlSheet = xlSheet;
                this.Estimates = LoadEstimates();
            }

            public override void BuildCorrelations()
            {
                BuildCorrelations_Input();
                BuildCorrelations_Periods();
            }

            private void BuildCorrelations_Input()
            {
                //Input correlation
                int maxDepth = (from Estimate est in this.Estimates select est.Level).Max();
                var correlTemp = BuildCorrelTemp(this.Estimates);
                if (Estimates.Any())
                    Estimates[0].xlCorrelCell_Inputs.EntireColumn.Clear();
                foreach (Estimate est in this.Estimates)
                {
                    PrintCorrel_Inputs(est, correlTemp);  //recursively build out children
                }
            }

            private void BuildCorrelations_Periods()
            {
                //Period correlation
                foreach (Estimate est in this.Estimates)
                {
                    //Save the existing values
                    if (est.xlCorrelCell_Periods != null)
                    {
                        est.xlCorrelCell_Periods.Clear();
                    }

                    PrintCorrel_Periods(est);
                }
            }

            private Dictionary<Tuple<UniqueID, UniqueID>, double> BuildCorrelTemp(List<IEstimate> Estimates)
            {
                var correlTemp = new Dictionary<Tuple<UniqueID, UniqueID>, double>();   //<ID, ID>, correl_value
                if (this.Estimates.Any())
                {
                    //Save off existing correlations
                    //Create a correl string from the column
                    foreach (Estimate estimate in this.Estimates)
                    {
                        if (estimate.SubEstimates.Count == 0)
                            continue;
                        Data.CorrelationString_Inputs correlString;
                        if (estimate.xlCorrelCell_Inputs.Value == null)        //No correlation string exists
                            correlString = Data.CorrelationString_Inputs.ConstructString(estimate.GetSubEstimateIDs(), this.xlSheet.Name);     //construct zero string
                        else
                            correlString = new Data.CorrelationString_Inputs(estimate.xlCorrelCell_Inputs.Value);       //construct from string
                        var correlMatrix = new Data.CorrelationMatrix(correlString);
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

            public override List<IEstimate> LoadEstimates()
            {
                List<IEstimate> returnList = new List<IEstimate>();
                Excel.Range lastCell = xlSheet.Cells[1000000, dc.Name_Offset].End[Excel.XlDirection.xlUp];
                Excel.Range firstCell = xlSheet.Cells[2, dc.Name_Offset];
                Excel.Range pullRange = xlSheet.Range[firstCell, lastCell];
                Excel.Range[] estRows = PullEstimates(pullRange.Address);
                int maxDepth = Convert.ToInt32((from Excel.Range row in estRows select row.Cells[1, LevelColumn].value).Max());

                for (int i = 1; i <= maxDepth; i++)
                {
                    Excel.Range[] topLevels = (from Excel.Range row in estRows where row.Cells[1, LevelColumn].value == i select row).ToArray<Excel.Range>();
                    for (int index = 0; index < topLevels.Count(); index++)
                    {
                        Estimate parentEstimate = new Estimate(topLevels[index].EntireRow);
                        parentEstimate.LoadSubEstimates();
                        returnList.Add(parentEstimate);
                    }
                }
                return returnList;
            }

            public override void PrintToSheet()
            {
                throw new NotImplementedException();
            }

            public override bool Validate()
            {
                throw new NotImplementedException();
            }
        }
    }

}