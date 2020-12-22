﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace CorrelationTest
{
    public enum CostItem
    {
        I,
        E,
        W,
        T
    }

    public abstract class CostSheet : Sheet, ICostSheet
    {
        protected DialogResult OverwriteRepeatedIDs { get; set; }
        protected DisplayCoords dc { get; set; }
        protected int LevelColumn { get; set; }

        public List<IEstimate> Estimates { get; set; }

        public virtual List<IEstimate> LoadEstimates()
        {
            throw new NotImplementedException();
        }

        public virtual object[] Get_xlFields()
        {
            throw new NotImplementedException();
        }

        public virtual void BuildCorrelations()
        {
            throw new NotImplementedException();
        }

        protected virtual void PrintCorrel_Inputs(Estimate estimate, Dictionary<Tuple<UniqueID, UniqueID>, double> inputTemp = null)
        {
            /*
             * This is being called when "Build" is run. 
             * 
             */
            if (estimate.SubEstimates.Count >= 2)
            {
                //DAVID
                //This has too many subestimates
                UniqueID[] subIDs = (from Estimate est in estimate.SubEstimates select est.uID).ToArray<UniqueID>();
                //check if any of the subestimates have NonZeroCorrel entries
                
                //This is sending in too many IDs
                Data.CorrelationString_Inputs correlationString_inputs = Data.CorrelationString_Inputs.ConstructString(subIDs, this.xlSheet.Name, inputTemp);
                correlationString_inputs.PrintToSheet(estimate.xlCorrelCell_Inputs);
            }
        }

        protected virtual void PrintCorrel_Periods(Estimate estimate, Dictionary<Tuple<PeriodID, PeriodID>, double> inputTemp = null)
        {
            /*
             * The print methods on the sheet object are there to compile a list of estimates
             * The print methods on the estimates should handle printing out correl strings
             * 
             * This should take a list of all estimates, recently built, cycle them, and call their print method to print correl strings (List<Estimate>)
             * The saved values should already be loaded into the estimates
             */
            //PeriodID[] periodIDs = (from Period prd in estimate.Periods select prd.pID).ToArray();
            //Data.CorrelationString_Periods correlationString_periods = Data.CorrelationString_Periods.ConstructString(periodIDs, this.xlSheet.Name);
            Data.CorrelationString_Periods correlationString_periods = Data.CorrelationString.Construct(estimate.xlCorrelCell_Periods.Value);
            correlationString_periods.PrintToSheet(estimate.xlCorrelCell_Periods);
        }

        public virtual Excel.Range[] PullEstimates(Excel.Range pullRange, CostItem costType) { return null; }

    }
}
