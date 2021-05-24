using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Winforms = System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;
using System.Threading;
//using System.Windows.Forms.DataVisualization.Charting;

namespace CorrelationTest
{
    public delegate Point ConvertToFormPoint(Point screenPoint);
    public delegate void PointAdder(DataPoint windowPoint);

    class DrawingTool
    {
        public Winforms.Button btn_ToolSwap { get; set; }
        public bool DrawPointsMode { get; set; } = true;

        private double XAxis_Min_Pixels { get; set; }
        private double YAxis_Min_Pixels { get; set; }
        private double XAxis_Max_Pixels { get; set; }
        private double YAxis_Max_Pixels { get; set; }

        private double XAxis_Min_Value { get; set; }
        private double YAxis_Min_Value { get; set; }
        private double XAxis_Max_Value { get; set; }
        private double YAxis_Max_Value { get; set; }

        private Dictionary<string, dynamic> Spacing { get; set; }

        public Series DrawSeries { get; set; } = new Series();  //The tool adds points to this series which is displayed in the DrawArea

        private Color CanvasColor { get; } = Color.FromArgb(195, 195, 195);
        private Color ExistingCanvasColor { get; set; }

        public Winforms.Cursor OverrideCursor { get; set; } = new Winforms.Cursor(new MemoryStream(Properties.Resources.CircleCursor1));
        public Winforms.Cursor ExistingCursor { get; set; }

        public System.Timers.Timer PaintTimer { get; set; } = new System.Timers.Timer();
        
        public Chart DrawOn { get; set; }
        public ChartArea DrawArea { get; set; }

        public DrawingTool(ref Chart DrawOn, ref ChartArea DrawArea, Dictionary<string, dynamic> Spacing)
        {
            this.DrawOn = DrawOn;
            this.DrawArea = DrawArea;
            this.Spacing = Spacing;
            this.ExistingCanvasColor = DrawArea.BackColor;
            DrawSeries.ChartType = SeriesChartType.Point;
            DrawSeries.Name = "DrawSeries";
            DrawSeries.Color = Color.FromArgb(255, 0, 162, 232);
            DrawSeries.MarkerBorderColor = Color.FromArgb(255, 0, 0, 0);
            DrawSeries.MarkerStyle = MarkerStyle.Circle;
            DrawSeries.MarkerSize = 8;
            DrawSeries.SmartLabelStyle.Enabled = false;
            this.PaintTimer.Interval = 200;     //ms
            PaintTimer.Enabled = false;
            PaintTimer.AutoReset = true;
            PaintTimer.Elapsed += AttemptPaint;
            if (!(from Series s in DrawOn.Series where s.Name == "DrawSeries" select s).Any())
            {
                DrawOn.Series.Add(DrawSeries);
            }
            
        }

        

        public void EnableSelectionMode()
        {
            DrawPointsMode = false;
            DrawArea.CursorX.IsUserEnabled = true;
            DrawArea.CursorY.IsUserEnabled = true;
            DrawArea.CursorX.IsUserSelectionEnabled = true;
            DrawArea.CursorY.IsUserSelectionEnabled = true;
            DrawArea.CursorX.Interval = 0.01;
            DrawArea.CursorY.Interval = 0.01;
            DrawArea.CursorX.AxisType = AxisType.Primary;
            DrawArea.CursorY.AxisType = AxisType.Primary;
            DrawArea.CursorX.LineColor = System.Drawing.Color.Transparent;
            DrawArea.CursorY.LineColor = System.Drawing.Color.Transparent;
            DrawArea.CursorX.SelectionColor = System.Drawing.Color.Transparent;
            DrawArea.CursorY.SelectionColor = System.Drawing.Color.Transparent;
            
        }

        public void EnableDrawPointsMode()
        {
            DrawPointsMode = true;
            DrawArea.CursorX.IsUserEnabled = false;
            DrawArea.CursorY.IsUserEnabled = false;
            DrawArea.CursorX.IsUserSelectionEnabled = false;
            DrawArea.CursorY.IsUserSelectionEnabled = false;
        }

        public void EnableCursor()
        {
            if(ExistingCursor == null)
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
            DrawOn.Series[0].Color = Color.FromArgb(0, DrawOn.Series[0].Color);
            DrawOn.Series[0].LabelBackColor = Color.FromArgb(0, DrawOn.Series[0].LabelBackColor);
            DrawOn.Series[0].LabelForeColor = Color.FromArgb(0, DrawOn.Series[0].LabelForeColor);
            DrawOn.Series[1].Color = Color.FromArgb(0, DrawOn.Series[1].Color);
            DrawOn.Series[1].LabelBackColor = Color.FromArgb(0, DrawOn.Series[1].LabelBackColor);
            DrawOn.Series[1].LabelForeColor = Color.FromArgb(0, DrawOn.Series[1].LabelForeColor);

            DrawSeries.Color = Color.FromArgb(255, DrawSeries.Color);
            DrawOn.Series["CorrelSeries"].Label = "";
        }

        public void ResetChartFormat()
        {
            //Reset the background color
            DrawOn.SendToBack();
            DrawArea.BackColor = ExistingCanvasColor;
            //Reset the marker colors
            DrawOn.Series[0].Color = Color.FromArgb(255, DrawOn.Series[0].Color);
            DrawOn.Series[0].LabelBackColor = Color.FromArgb(255, DrawOn.Series[0].LabelBackColor);
            DrawOn.Series[0].LabelForeColor = Color.FromArgb(255, DrawOn.Series[0].LabelForeColor);
            DrawOn.Series[1].Color = Color.FromArgb(255, DrawOn.Series[1].Color);
            DrawOn.Series[1].LabelBackColor = Color.FromArgb(255, DrawOn.Series[1].LabelBackColor);
            DrawOn.Series[1].LabelForeColor = Color.FromArgb(255, DrawOn.Series[1].LabelForeColor);

            DrawSeries.Color = Color.FromArgb(0, DrawSeries.Color);

            DrawArea.CursorX.IsUserEnabled = false;
            //DrawArea.CursorX.IsUserSelectionEnabled = false;
            DrawArea.CursorY.IsUserEnabled = false;
            //DrawArea.CursorY.IsUserSelectionEnabled = false;
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

        public void PaintPoint()
        {
            int x = System.Windows.Forms.Cursor.Position.X;
            int y = System.Windows.Forms.Cursor.Position.Y;
            Accord.Statistics.Distributions.Univariate.NormalDistribution offsetDist = new Accord.Statistics.Distributions.Univariate.NormalDistribution(0, 9);
            Random rando = new Random();
            int x_offset = Convert.ToInt32(offsetDist.InverseDistributionFunction(rando.NextDouble()));
            int y_offset = Convert.ToInt32(offsetDist.InverseDistributionFunction(rando.NextDouble()));
            Point screenPoint = new Point(x + x_offset, y + y_offset);
            Point newPoint;

            if (DrawOn.InvokeRequired)
            {
                ConvertToFormPoint cfp = DrawOn.PointToClient;
                newPoint = (Point)DrawOn.Invoke(cfp, screenPoint);
            }
            else
            {
                newPoint = DrawOn.PointToClient(screenPoint);
            }
            AddPoint(newPoint);
        }

        private void AttemptPaint(object sender, EventArgs e)   //This event fires from the timer elapsed event.
        {            
            PaintPoint();
        }

        public decimal GetCorrelationFromPoints()
        {
            if(this.DrawSeries.Points.Count() < 3)
            {
                return -2;      //Not enough points
            }
            else
            {
                double[,] matrix = new double[DrawSeries.Points.Count(), 2];
                var xVals = (from DataPoint dp in DrawSeries.Points select dp.XValue).ToArray();
                var yVals = (from DataPoint dp in DrawSeries.Points select dp.YValues.First()).ToArray();
                for(int i = 0; i < matrix.GetLength(0); i++)
                {
                    matrix[i, 0] = xVals[i];
                    matrix[i, 1] = yVals[i];
                }
                double[,] corMatrix = Accord.Statistics.Measures.Correlation(matrix);
                return Convert.ToDecimal(corMatrix[0, 1]);
            }
        }

        public void AddPoint(Point newPoint)
        {
            //Need to check boundaries still. Points aren't showing outside the innerplot area, but are being added to the series regardless
            double leftBound = Spacing["chartInnerPlot_Abs_Left"];
            double topBound = Spacing["chartInnerPlot_Abs_Top"];
            double rightBound = Spacing["chartInnerPlot_Abs_Right"];
            double bottomBound = Spacing["chartInnerPlot_Abs_Bottom"]; ;
            if(newPoint.X >= leftBound && newPoint.X <= rightBound && newPoint.Y <= bottomBound && newPoint.Y >= topBound)
            {

                DataPoint newDataPoint = ConvertPointToDataPoint(newPoint);
                if (DrawOn.InvokeRequired)
                {
                    PointAdder pa = DrawSeries.Points.Add;
                    DrawOn.Invoke(pa, newDataPoint);
                }
                else
                {
                    DrawSeries.Points.Add(newDataPoint);
                }
            }

            ////PLAN: NEED TO CONVERT the Point's coords to axis values (axisToValue inside paint event?), create a DataPoint off the derived x & y
            //Check if you are within the InnerPlot area - if so, add and return true. If not, return false.
        }

        
    }
}
