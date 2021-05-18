using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CorrelationTest
{
    class DrawnCorrelation
    {
        public List<Point> Points { get; set; } = new List<Point>();

        public DrawnCorrelation() { }

        

        public Point[] GetPoints()
        {
            return Points.ToArray();
        }

        public double GetSlope()
        {
            //Take the points and pull an OLS slope out.
            double[] xVals = new double[Points.Count()];
            double[] yVals = new double[Points.Count()];
            xVals = (from Point pt in Points select Convert.ToDouble(pt.X)).ToArray();
            yVals = (from Point pt in Points select Convert.ToDouble(-1 * pt.Y)).ToArray();

            var slr = new Accord.Statistics.Models.Regression.Linear.SimpleLinearRegression();
            var ols = new Accord.Statistics.Models.Regression.Linear.OrdinaryLeastSquares();
            slr = ols.Learn(xVals, yVals);
            return slr.Slope;
        }

        public double GetCorrelation()
        {
            double[] xVals = new double[Points.Count()];
            double[] yVals = new double[Points.Count()];
            xVals = (from Point pt in Points select Convert.ToDouble(pt.X)).ToArray();
            yVals = (from Point pt in Points select Convert.ToDouble(pt.Y)).ToArray();

            var slr = new Accord.Statistics.Models.Regression.Linear.SimpleLinearRegression();
            var ols = new Accord.Statistics.Models.Regression.Linear.OrdinaryLeastSquares();
            slr = ols.Learn(xVals, yVals);
            return Math.Sqrt(slr.CoefficientOfDetermination(xVals, yVals));
        }
    }
}
