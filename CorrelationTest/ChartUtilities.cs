using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms.DataVisualization.Charting;

namespace CorrelationTest
{
    public static class ChartUtilities
    {

        public static SelectedPoint SelectDataPointNearToXY(double x, double y, Series series)
        {
            if (series.Points.Count() == 0)
                return null;

            var distances = (from DataPoint dp in series.Points select new Tuple<DataPoint, double>(dp, GetDistance(dp, x, y))).OrderBy(t => t.Item2);
            DataPoint closestDataPoint = distances.First().Item1;
            double nearestDistance = distances.First().Item2;
            if (nearestDistance <= 0.15)
                return new SelectedPoint(closestDataPoint, series);
            else
                return null;
        }


        private static double GetDistance(DataPoint dp, double x, double y)
        {
            double dp_x = dp.XValue;
            double dp_y = dp.YValues.First();

            double distance_x = x - dp_x;
            double distance_y = y - dp_y;

            return Math.Sqrt(distance_x * distance_x + distance_y * distance_y); //Pythagorean theorem!
        }
    }
}
