using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CorrelationTest
{
    public class PeriodID : UniqueID
    {
        public int PeriodTag { get; set; }
        public PeriodID(UniqueID uID, string start_date)
        {
            this.ID = $"{uID.ID}_{start_date}";
        }

        public static PeriodID[] GeneratePeriodIDs(UniqueID uid, int numOfPeriods)
        {
            PeriodID[] pids = new PeriodID[numOfPeriods];
            for(int i = 0; i < numOfPeriods; i++)
            {
                pids[i] = new PeriodID(uid, (i+1).ToString());
            }
            return pids;
        }
    }
}
