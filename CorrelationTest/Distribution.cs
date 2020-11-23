using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Accord.Statistics.Distributions.Univariate;
using Accord.Statistics.Distributions;

namespace CorrelationTest
{
    public static class DistributionType
    {
        public const string Normal = "Normal";
        public const string Triangular = "Triangular";
        public const string Lognormal = "Lognormal";
        public const string Beta = "Beta";
    }

    public class Distribution
    {
        public string Name { get; set; }
        public IUnivariateDistribution DistributionObj { get; set; }
        public string DistributionString { get; set; }
        
        public Distribution(string distributionString)
        {
            this.DistributionObj = BuildDistribution(ParseString(distributionString));
        }

        public Distribution(Dictionary<string, object> distParameters)
        {
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
    }
}
