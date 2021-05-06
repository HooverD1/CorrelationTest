using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CorrelationTest
{
    class LineSegment
    {
        public List<Point> Points { get; set; } = new List<Point>();

        public LineSegment() { }
        public LineSegment(Point firstPoint)
        {
            this.Points.Add(firstPoint);
        }

        public void AddPoint(Point newPoint)
        {
            this.Points.Add(newPoint);
        }

        public Point[] GetPoints()
        {
            return Points.ToArray();
        }
    }
}
