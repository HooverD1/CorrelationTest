using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace CorrelationTest
{
    public class CustomDistribution : IEstimateDistribution
    {
        //This class will be its own distribution object

        public string Name { get; set; }
        public string DistributionString { get; set; }
        public Dictionary<string, object> DistributionParameters { get; set; }

        private CustomDistribution() { } //Default

        //EXPAND
        public static IEstimateDistribution ConstructForExpansion(Excel.Range xlRow, DisplayCoords specs, CorrelationType correlType)
        {
            //Pull the name and parameters to move it to the correlation sheet, but do not do any calculation
            IEstimateDistribution returnObject = new CustomDistribution();

            returnObject.Name = "Custom";
            if (correlType == CorrelationType.Cost || correlType == CorrelationType.Duration)
                returnObject.DistributionString = xlRow.Cells[1, specs.Distribution_Offset].Value;
            else if (correlType == CorrelationType.Phasing)
                returnObject.DistributionString = xlRow.Cells[1, specs.PhasingCorrel_Offset].Value;
            else
                throw new Exception("Unexpected Correlation Type");
            returnObject.DistributionParameters = new Dictionary<string, object>();
            returnObject.DistributionParameters.Add("Param1", 1);
            returnObject.DistributionParameters.Add("Param2", 2);
            returnObject.DistributionParameters.Add("Param3", 3);
            
            return returnObject;
        }

        //VISUALIZATION
        public static IEstimateDistribution ConstructForVisualization(Excel.Range xlRow, Sheets.CorrelationSheet cs)
        {
            //Need to know which item on the correlsheet we're talking about (the xlRow of the selection)
            //Need to know the xlSheet and specs off the CorrelationSheet (pass the sheet object)
            IEstimateDistribution returnObject = new CustomDistribution();

            returnObject.Name = xlRow.Cells[1, cs.Specs.DistributionCoords.Item2];
            returnObject.DistributionString = xlRow.Cells[1, cs.Specs.DistributionCoords.Item2].Value;
            returnObject.DistributionParameters = ParseStringIntoParameters(returnObject.DistributionString);

            return returnObject;
        }

        //

        public double GetInverse(double percentile)
        {
            throw new NotImplementedException();
        }

        private static Dictionary<string, object> ParseStringIntoParameters(string distributionString)
        {
            throw new NotImplementedException();
        }
    }
}
