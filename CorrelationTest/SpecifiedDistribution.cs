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
        private IUnivariateDistribution DistributionObj { get; set; }        //Specified distributions will contain an accord object

        public string Name { get; set; }
        public string DistributionString { get; set; }
        public Dictionary<string, object> DistributionParameters { get; set; }
        
        public SpecifiedDistribution(string distributionString)
        {            
            this.DistributionParameters = ParseString(distributionString);
            this.Name = DistributionParameters["Type"].ToString();
            this.DistributionObj = BuildDistribution(this.DistributionParameters);
        }

        public SpecifiedDistribution(Dictionary<string, object> distParameters)
        {
            //Switch between standard distributions and the custom aggregate one based on inputs
            this.DistributionParameters = distParameters;
            this.Name = distParameters["Type"].ToString();
            this.DistributionObj = BuildDistribution(distParameters);
        }

        public Dictionary<string, object> ParseString(string distributionString)
        {
            Dictionary<string, object> stringItems = new Dictionary<string, object>();
            string[] splitString = distributionString.Split(',');
            stringItems.Add("Type", splitString[0]);
            for (int i = 1; i < splitString.Length; i++)
                stringItems.Add($"Param{i}", splitString[i]);
            return stringItems;
        }

        public double GetInverse(double percentile)
        {
            return DistributionObj.InverseDistributionFunction(percentile);
        }

        private IUnivariateDistribution BuildDistribution(Dictionary<string, object> distParameters)
        {
            switch (distParameters["Type"])
            {
                case (DistributionType.Triangular): //Min, Max, Mode
                    return new TriangularDistribution(Convert.ToDouble(distParameters["Param1"]), Convert.ToDouble(distParameters["Param2"]), Convert.ToDouble(distParameters["Param3"]));
                case (DistributionType.Normal):
                    return new NormalDistribution(Convert.ToDouble(distParameters["Param1"]), Convert.ToDouble(distParameters["Param2"]));      //mean, stdev
                case (DistributionType.Lognormal):
                    return new LognormalDistribution(Convert.ToDouble(distParameters["Param1"]), Convert.ToDouble(distParameters["Param2"]));   //shape, location
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
