using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Winforms = System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;
//using System.Windows.Forms.DataVisualization.Charting;

namespace CorrelationTest
{
    class DrawingTool
    {
        private double XAxis_Min_Pixels { get; set; }
        private double YAxis_Min_Pixels { get; set; }
        private double XAxis_Max_Pixels { get; set; }
        private double YAxis_Max_Pixels { get; set; }

        private double XAxis_Min_Value { get; set; }
        private double YAxis_Min_Value { get; set; }
        private double XAxis_Max_Value { get; set; }
        private double YAxis_Max_Value { get; set; }

        private bool leftRestricted = false;
        private bool rightRestricted = false;

        public Series DrawSeries { get; set; } = new Series();  //The tool adds points to this series which is displayed in the DrawArea

        private Color CanvasColor { get; } = Color.FromArgb(195, 195, 195);
        private Color ExistingCanvasColor { get; set; }

        public Winforms.Cursor OverrideCursor { get; set; } = new Winforms.Cursor(new MemoryStream(Properties.Resources.CircleCursor1));
        public Winforms.Cursor ExistingCursor { get; set; }
        
        public Chart DrawOn { get; set; }
        public ChartArea DrawArea { get; set; }

        public DrawingTool(ref Chart DrawOn, ref ChartArea DrawArea)
        {
            this.DrawOn = DrawOn;
            this.DrawArea = DrawArea;
            this.ExistingCanvasColor = DrawArea.BackColor;
            DrawSeries.ChartType = SeriesChartType.Point;
            DrawSeries.Name = "DrawSeries";
            DrawSeries.Color = Color.FromArgb(255, 1, 1, 1);
            
            if (!(from Series s in DrawOn.Series where s.Name == "DrawSeries" select s).Any())
            {
                DrawOn.Series.Add(DrawSeries);
            }
        }

        public void EnableCursor()
        {
            ExistingCursor = Winforms.Cursor.Current;
            DrawOn.Cursor = this.OverrideCursor;

        }
        public void ResetCursor()
        {
            DrawOn.Cursor = ExistingCursor;
        }

        public void FormatChartForDrawing()
        {
            //Setup CorrelScatter for drawing

            //Set the background to the canvas color
            DrawOn.BringToFront();
            DrawArea.BackColor = CanvasColor;
            //Hide the markers
            foreach (Series s in DrawOn.Series)
            {
                s.Color = Color.FromArgb(0, s.Color);
            }
            DrawSeries.Color = Color.FromArgb(255, DrawSeries.Color);
        }

        public void ResetChartFormat()
        {
            //Reset the background color
            DrawOn.SendToBack();
            DrawArea.BackColor = ExistingCanvasColor;
            //Reset the marker colors
            foreach(Series s in DrawOn.Series)
            {
                s.Color = Color.FromArgb(255, s.Color);
            }
            DrawSeries.Color = Color.FromArgb(0, DrawSeries.Color);
        }

        public void GetXAxisMinMax()
        {
            XAxis_Min_Value = DrawArea.AxisX.Minimum;
            XAxis_Min_Pixels = DrawArea.AxisX.ValueToPixelPosition(DrawArea.AxisX.Minimum);
            XAxis_Max_Value = DrawArea.AxisX.Maximum;
            XAxis_Max_Pixels = DrawArea.AxisX.ValueToPixelPosition(DrawArea.AxisX.Maximum);
        }

        public void GetYAxisMinMax()
        {
            YAxis_Min_Value = DrawArea.AxisY.Minimum;
            YAxis_Min_Pixels = DrawArea.AxisY.ValueToPixelPosition(DrawArea.AxisY.Minimum);
            YAxis_Max_Value = DrawArea.AxisY.Maximum;
            YAxis_Max_Pixels = DrawArea.AxisY.ValueToPixelPosition(DrawArea.AxisY.Maximum);
        }

        private DataPoint ConvertPointToDataPoint(Point point)
        {
            double x_depth = (double)point.X - XAxis_Min_Pixels;
            double x_percentage = x_depth / (XAxis_Max_Pixels - XAxis_Min_Pixels);
            double y_depth = YAxis_Min_Pixels - (double)point.Y;    //Distance up from bottom
            double y_percentage = y_depth / (YAxis_Min_Pixels - YAxis_Max_Pixels);

            double x_value = x_percentage * (DrawArea.AxisX.Maximum - DrawArea.AxisX.Minimum) + DrawArea.AxisX.Minimum;
            double y_value = y_percentage * (DrawArea.AxisY.Maximum - DrawArea.AxisY.Minimum) + DrawArea.AxisY.Minimum;

            return new DataPoint(x_value, y_value);
        }

        public bool AddPoint(Point newPoint)
        {
            DataPoint newDataPoint = ConvertPointToDataPoint(newPoint);
            ////PLAN: NEED TO CONVERT the Point's coords to axis values (axisToValue inside paint event?), create a DataPoint off the derived x & y
            //Check if you are within the InnerPlot area - if so, add and return true. If not, return false.
            DrawSeries.Points.Add(newDataPoint);
            return true;
            //if (DrawSeries.Points.Count == 0)   //If this is the first point, it can go anywhere
            //{

            //    return true;
            //}
            //else if (DrawSeries.Points.Count == 1)  //Second point chooses which direction you are drawing
            //{
            //    //Check if X is equal.
            //    if (newDataPoint.XValue == DrawSeries.Points.First().XValue)
            //    {
            //        return false;
            //    }
            //    else if (newDataPoint.XValue > DrawSeries.Points.First().XValue)  //moving right
            //    {
            //        leftRestricted = true;
            //        DrawSeries.Points.Add(newDataPoint);
            //        return true;
            //    }
            //    else //if(newPoint.X < Points.First().X)  //moving left
            //    {
            //        rightRestricted = true;
            //        DrawSeries.Points.Add(newDataPoint);
            //        return true;
            //    }
            //}
            //else   //Additional points must go in the same direction
            //{
            //    //Check if the point violates the restriction
            //    if (newDataPoint.XValue == DrawSeries.Points.Last().XValue)
            //    {
            //        return false;
            //    }
            //    else if (newDataPoint.XValue > DrawSeries.Points.Last().XValue && rightRestricted == false)
            //    {
            //        DrawSeries.Points.Add(newDataPoint);
            //        return true;
            //    }
            //    else if (newDataPoint.XValue < DrawSeries.Points.Last().XValue && leftRestricted == false)
            //    {
            //        DrawSeries.Points.Add(newDataPoint);
            //        return true;
            //    }
            //    else
            //    {
            //        return false;
            //    }
            //}

        }
    }
}
