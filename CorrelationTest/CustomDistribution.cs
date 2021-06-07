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
        public static IEstimateDistribution ConstructForVisualization(Excel.Range xlSelection, Sheets.CorrelationSheet cs)
        {
            //Load the distribution object off a given row of the Correlation Sheet such that it can be leveraged for display purposes
            IEstimateDistribution returnObject = new CustomDistribution();

            returnObject.Name = Convert.ToString(xlSelection.EntireRow.Cells[1, cs.Specs.DistributionCoords.Item2].value);
            returnObject.DistributionString = Convert.ToString(xlSelection.EntireRow.Cells[1, cs.Specs.DistributionCoords.Item2].Value);
            returnObject.DistributionParameters = ParseStringIntoParameters(returnObject.DistributionString);

            return returnObject;
        }

        public double GetMaximum()
        {
            throw new NotImplementedException();
        }

        public double GetMinimum()
        {
            throw new NotImplementedException();
        }

        public double GetInverse(double percentile)
        {
            throw new NotImplementedException();
        }

        public double GetPDF_Value(double xValue)
        {
            throw new NotImplementedException();
        }

        public double GetPDF_MaxHeight()
        {
            //Find the maximum Y value for the PDF function -- implemented as an approx via recursive search function in SpecifiedDistribution
            throw new NotImplementedException();
        }

        public double GetCDF_Value(double xValue)
        {
            throw new NotImplementedException();
        }

        public double GetMean()
        {
            throw new NotImplementedException();
        }

        public double GetStdev()
        {
            throw new NotImplementedException();
        }

        private static Dictionary<string, object> ParseStringIntoParameters(string distributionString)
        {
            string[] distributionStringValues = distributionString.Split(',');
            Dictionary<string, object> parameters = new Dictionary<string, object>();
            parameters.Add("Type", distributionStringValues[0]);
            if (Double.TryParse(distributionStringValues[1], out double mean))
                parameters.Add("Mean", mean);
            if (Double.TryParse(distributionStringValues[2], out double stdev))
                parameters.Add("Stdev", stdev);

            var lookupDictionary = new Dictionary<string, object>();
            
            //Add look-up values to the dictionary here

            parameters.Add("LookUp", lookupDictionary);

            return parameters;
        }
    }
}
