using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace CorrelationTest
{
    public static class Distribution
    {
        public static IEstimateDistribution ConstructForExpansion(Excel.Range xlRow, CorrelationType correlType)
        {
            DisplayCoords specs = DisplayCoords.ConstructDisplayCoords(ExtensionMethods.GetSheetType(xlRow.Worksheet));
            //Need to get the name from the xlRow and return the appropriate type
            string distName = xlRow.Cells[1, specs.Distribution_Offset].value;
            switch (distName)
            {
                case "Custom":
                    return CustomDistribution.ConstructForExpansion(xlRow, specs, correlType);
                case "Normal":
                    return SpecifiedDistribution.ConstructForExpansion(xlRow, correlType);
                case "Triangular":
                    return SpecifiedDistribution.ConstructForExpansion(xlRow, correlType);
                case "Lognormal":
                    return SpecifiedDistribution.ConstructForExpansion(xlRow, correlType);
                case null:
                    return null;
                default:
                    throw new Exception("Unknown distribution type");
            }
        }

        public static IEstimateDistribution ConstructForVisualization(Excel.Range xlRow, Sheets.CorrelationSheet cs)
        {
            //Need to get the name from the xlRow and return the appropriate type
            string distString = xlRow.Cells[1, cs.Specs.DistributionCoords.Item2].value;
            if (distString is null)
                return null;
            string[] distSplit = distString.Split(',');
            switch (distSplit[0])
            {
                case "Custom":
                    return CustomDistribution.ConstructForVisualization(xlRow, cs);
                case "Normal":
                    return SpecifiedDistribution.ConstructForVisualization(xlRow, cs);
                default:
                    throw new Exception("Unknown distribution type");
            }
        }

    }
}
