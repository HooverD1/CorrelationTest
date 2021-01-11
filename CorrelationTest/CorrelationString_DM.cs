using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CorrelationTest
{
    namespace Data
    {
        public class CorrelationString_DM : CorrelationString
        {
            public CorrelationString_DM(string correlStringValue)
            {
                this.Value = correlStringValue;
            }

            public static bool Validate()
            {
                return true;
            }
        }
    }
}
