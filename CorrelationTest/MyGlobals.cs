using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CorrelationTest
{
    public static class MyGlobals
    {
        //THIS SHOULD BE FALSE WHEN BUILT FOR RELEASE
        public static bool DebugMode { get; set; } = true;
        
    }
}
