using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CorrelationTest
{
    namespace Data
    {
        public class CorrelationString_DT : CorrelationString
        {
            public CorrelationString_DT(string correlStringValue)
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
