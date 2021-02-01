﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace CorrelationTest
{
    public interface IJointEstimate : IHasSubs
    {
        CostEstimate costEstimate { get; set; }
        ScheduleEstimate scheduleEstimate { get; set; }

        CostEstimate ConstructCostSubEstimate(Excel.Range xlRow, CostSheet ContainingSheetObject);
        ScheduleEstimate ConstructScheduleSubEstimate(Excel.Range xlRow, CostSheet ContainingSheetObject);
    }
}
