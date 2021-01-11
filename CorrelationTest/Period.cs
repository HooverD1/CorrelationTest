using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace CorrelationTest
{
    public class Period
    {
        public int Number { get; set; }
        public string Start_Date { get; set; }
        public PeriodID pID { get; set; }       //ID of the period
        public UniqueID uID { get; set; }       //ID of the parent
        public double Dollars { get; set; }
        public Excel.Range xlCell { get; set; }

        public Period(UniqueID uID, string start_date, double Dollars = 0)
        {
            this.Start_Date = start_date;
            this.uID = uID;
            this.pID = new PeriodID(uID, Number);
            this.Dollars = Dollars;
        }

    }
}
