﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace CorrelationTest
{
    public class Input_Item : Item, ISub, IHasPhasingSubs
    {
        public Excel.Range xlDollarCell { get; set; }
        public DisplayCoords dispCoords { get; set; }
        public int PeriodCount { get; set; }
        public Period[] Periods { get; set; }
        public Distribution ItemDistribution { get; set; }
        public Dictionary<string, object> DistributionParameters { get; set; }
        public Dictionary<Estimate_Item, double> CorrelPairs { get; set; }

        public Input_Item(Excel.Range xlRow, CostSheet ContainingSheetObject) : base(xlRow, ContainingSheetObject)
        {
            LoadPeriods();
        }

        public void LoadPeriods()
        {
            this.Periods = GetPeriods();
        }
        private Period[] GetPeriods()       //should these be constructed as a static under Period?
        {
            Period[] periods = new Period[5];
            for (int i = 0; i < periods.Length; i++)
                periods[i] = new Period(uID, i + 1);
            return periods;
        }

        public void LoadUID()
        {
            this.uID = GetUID();
        }
        private UniqueID GetUID()
        {
            throw new NotImplementedException();
        }

        public void PrintPhasingCorrelString()
        {
            Data.CorrelationString phString = Data.CorrelationString.Construct(this, Data.CorrelStringType.PhasingTriple);
            if (phString != null)
                phString.PrintToSheet(xlCorrelCell_Periods);
        }
    }
}
