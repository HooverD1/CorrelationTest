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
        private bool leftRestricted = false;
        private bool rightRestricted = false;

        public LineSegment() { }

        public bool AddPoint(Point newPoint)
        {
            //Check if its crossing back over X values.
            if(Points.Count == 0)   //If this is the first point, it can go anywhere
            {
                this.Points.Add(newPoint);
                return true;
            }
            else if(Points.Count == 1)  //Second point chooses which direction you are drawing
            {
                //Check if X is equal.
                if(newPoint.X == Points.First().X)
                {
                    return false;
                }
                else if(newPoint.X > Points.First().X)  //moving right
                {
                    leftRestricted = true;
                    this.Points.Add(newPoint);
                    return true;
                }
                else //if(newPoint.X < Points.First().X)  //moving left
                {
                    rightRestricted = true;
                    this.Points.Add(newPoint);
                    return true;
                }
            }
            else   //Additional points must go in the same direction
            {
                //Check if the point violates the restriction
                if(newPoint.X == Points.Last().X)
                {
                    return false;
                }
                else if(newPoint.X > Points.Last().X && rightRestricted == false)
                {
                    this.Points.Add(newPoint);
                    return true;
                }
                else if(newPoint.X < Points.Last().X && leftRestricted == false)
                {
                    this.Points.Add(newPoint);
                    return true;
                }
                else
                {
                    return false;
                }
            }

        }

        public Point[] GetPoints()
        {
            return Points.ToArray();
        }
    }
}
