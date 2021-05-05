using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Accord.Statistics.Distributions.Univariate;
using Accord.Statistics.Distributions;
using Excel = Microsoft.Office.Interop.Excel;

namespace CorrelationTest
{
    public static class DistributionType
    {
        public const string Normal = "Normal";
        public const string Triangular = "Triangular";
        public const string Lognormal = "Lognormal";
        public const string Beta = "Beta";
    }

    public class SpecifiedDistribution : IEstimateDistribution
    {
        //This class will wrap an Accord object

        private IUnivariateDistribution Distribution { get; set; }        //Specified distributions will contain an accord object

        public string Name { get; set; }
        public string DistributionString { get; set; }
        public Dictionary<string, object> DistributionParameters { get; set; }

        private SpecifiedDistribution() { }
        
        //public SpecifiedDistribution(string distributionString)
        //{            
        //    this.DistributionParameters = ParseStringIntoParameters(distributionString);
        //    this.Name = DistributionParameters["Type"].ToString();
        //    this.Distribution = BuildDistribution(this.DistributionParameters);
        //}

        //public SpecifiedDistribution(Dictionary<string, object> distParameters)
        //{
        //    //Switch between standard distributions and the custom aggregate one based on inputs
        //    this.DistributionParameters = distParameters;
        //    this.Name = distParameters["Type"].ToString();
        //    this.Distribution = BuildDistribution(distParameters);
        //}

        //EXPAND
        public static IEstimateDistribution ConstructForExpansion(Excel.Range xlRow, CorrelationType correlType)
        {
            DisplayCoords specs = DisplayCoords.ConstructDisplayCoords(ExtensionMethods.GetSheetType(xlRow.Worksheet));
            //Pull the name and parameters to move it to the correlation sheet, but do not do any calculation
            IEstimateDistribution returnObject = new SpecifiedDistribution();

            returnObject.Name = "Custom";
            if (correlType == CorrelationType.Cost || correlType == CorrelationType.Duration)
                returnObject.DistributionString = xlRow.Cells[1, specs.Distribution_Offset].Value;
            else if (correlType == CorrelationType.Phasing)
                returnObject.DistributionString = xlRow.Cells[1, specs.Phasing_Offset].Value;
            else
                throw new Exception("Unexpected correlation type");

            returnObject.DistributionParameters = new Dictionary<string, object>();
            returnObject.DistributionParameters.Add("Param1", 1);
            returnObject.DistributionParameters.Add("Param2", 2);
            returnObject.DistributionParameters.Add("Param3", 3);

            return returnObject;
        }

        //VISUALIZATION
        public static IEstimateDistribution ConstructForVisualization(Excel.Range xlSelection, Sheets.CorrelationSheet cs)
        {
            //Need to know which item on the correlsheet we're talking about (the xlRow of the selection)
            //Need to know the xlSheet and specs off the CorrelationSheet (pass the sheet object)
            SpecifiedDistribution returnObject = new SpecifiedDistribution();
            string distString = xlSelection.EntireRow.Cells[1, cs.Specs.DistributionCoords.Item2].value;
            returnObject.Name = distString.Split(',')[0];
            returnObject.DistributionString = distString;
            returnObject.DistributionParameters = ParseStringIntoParameters(returnObject.DistributionString);
            returnObject.Distribution = BuildDistribution(returnObject.DistributionParameters);
            return returnObject;
        }

        private static Dictionary<string, object> ParseStringIntoParameters(string distributionString)
        {
            Dictionary<string, object> stringItems = new Dictionary<string, object>();
            string[] splitString = distributionString.Split(',');
            stringItems.Add("Type", splitString[0]);
            switch (stringItems["Type"])
            {
                case "Normal":
                    if (splitString.Length != 3)
                        throw new Exception("Malformed distribution string");
                    stringItems.Add("Mean", splitString[1]);
                    stringItems.Add("Stdev", splitString[2]);
                    break;
                case "Lognormal":
                    if (splitString.Length != 3)
                        throw new Exception("Malformed distribution string");
                    stringItems.Add("Mean", splitString[1]);
                    stringItems.Add("Stdev", splitString[2]);
                    break;
                case "Triangular":
                    if (splitString.Length != 4)
                        throw new Exception("Malformed distribution string");
                    stringItems.Add("Min", splitString[1]);
                    stringItems.Add("Max", splitString[2]);
                    stringItems.Add("Mode", splitString[3]);
                    break;
                case "Beta":
                    if (splitString.Length != 5)
                        throw new Exception("Malformed distribution string");
                    stringItems.Add("Mean", splitString[1]);
                    stringItems.Add("Stdev", splitString[2]);
                    stringItems.Add("Alpha", splitString[3]);
                    stringItems.Add("Beta", splitString[4]);
                    break;
                default:
                    throw new Exception("Unknown distribution type");
            }
            return stringItems;
        }

        public double GetInverse(double percentile)
        {
            return Distribution.InverseDistributionFunction(percentile);
        }

        public double GetMaximum_X()
        {
            double stdev = Math.Sqrt(Distribution.Variance);
            return Distribution.Mean + stdev * 4;
        }

        public double GetMinimum_X()
        {
            if (Distribution is LognormalDistribution)
            {
                return 0;
            }
            else
            {
                double stdev = Math.Sqrt(Distribution.Variance);
                return Distribution.Mean - stdev * 4;
            }
        }

        public double GetMaximum_Y()
        {
            throw new NotImplementedException();
        }

        public double GetMinimum_Y()
        {
            throw new NotImplementedException();
        }

        public double GetPDF_Value(double xValue)
        {
            //Get the Y value from the X value
            if (Distribution is LognormalDistribution && xValue == 0)
                return 0;
            return Distribution.ProbabilityFunction(xValue);
        }

        public double GetCDF_Value(double xValue)
        {
            //Get the Y value from the X value
            return Distribution.DistributionFunction(xValue);
        }

        private static IUnivariateDistribution BuildDistribution(Dictionary<string, object> distParameters)
        {
            switch (distParameters["Type"])
            {
                case (DistributionType.Triangular): //Min, Max, Mode
                    return new TriangularDistribution(Convert.ToDouble(distParameters["Min"]), Convert.ToDouble(distParameters["Max"]), Convert.ToDouble(distParameters["Mode"]));
                case (DistributionType.Normal):
                    return new NormalDistribution(Convert.ToDouble(distParameters["Mean"]), Convert.ToDouble(distParameters["Stdev"]));      //mean, stdev
                case (DistributionType.Lognormal):
                    return new LognormalDistribution(Convert.ToDouble(distParameters["Mean"]), Convert.ToDouble(distParameters["Stdev"]));   //shape, location
                case (DistributionType.Beta):
                    return new BetaDistribution(Convert.ToDouble(distParameters["Param1"]), Convert.ToDouble(distParameters["Param2"]));        //alpha, beta
                default:
                    return null;
            }
        }

        public static string GetDistributionStringFromRange(Excel.Range xlDistributionCell)
        {
            Excel.Range xlDistributionRange = xlDistributionCell.Resize[1, 5];
            object[,] distributionValues = xlDistributionRange.Value;
            StringBuilder distributionString = new StringBuilder();
            for (int i = 1; i <= 5; i++)
            {
                distributionString.Append(distributionValues[1,i]);
                distributionString.Append(",");
            }
            distributionString.Remove(distributionString.Length - 1, 1);    //remove the final char
            return distributionString.ToString();
        }
    }
}
