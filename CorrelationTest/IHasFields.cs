using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace CorrelationTest
{
    public interface IHas_xlFields
    {
        Excel.Worksheet xlSheet { get; set; }
        object[] Get_xlFields();

    }
}
