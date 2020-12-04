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
        public PeriodID(UniqueID uID, int period) : base(uID.ID)
        {
            this.PeriodTag = period;
            ID = $"{uID.ID}{Delimiter}{period}";
        }
    }
}
