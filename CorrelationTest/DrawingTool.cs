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
using System.Windows.Forms;
//using System.Windows.Forms.DataVisualization.Charting;

namespace CorrelationTest
{
    public delegate Point ConvertToFormPoint(Point screenPoint);
    public delegate void PointAdder(DataPoint windowPoint);

    class DrawingTool
    {
        public Winforms.Button btn_ToolSwap { get; set; }
        public Winforms.Button btn_ClearPoints { get; set; }
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

        private const double PaintTimerDefaultRate = 150;       //The slowest the points ever drop is 5x per second.
        public System.Timers.Timer PaintTimer { get; set; } = new System.Timers.Timer();
        public int PointsPerTimer { get; set; } = 1; //default to 1
        
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
            this.PaintTimer.Interval = 200;     //default
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
            DrawArea.BackColor = CanvasColor;
            //Hide the markers
            DrawOn.Series["CorrelSeries"].Color = Color.FromArgb(0, DrawOn.Series["CorrelSeries"].Color);
            DrawOn.Series["CorrelSeries"].LabelBackColor = Color.FromArgb(0, DrawOn.Series["CorrelSeries"].LabelBackColor);
            DrawOn.Series["CorrelSeries"].LabelForeColor = Color.FromArgb(0, DrawOn.Series["CorrelSeries"].LabelForeColor);
            DrawOn.Series["MeanMarker"].Color = Color.FromArgb(0, DrawOn.Series["MeanMarker"].Color);
            DrawOn.Series["MeanMarker"].LabelBackColor = Color.FromArgb(0, DrawOn.Series["MeanMarker"].LabelBackColor);
            DrawOn.Series["MeanMarker"].LabelForeColor = Color.FromArgb(0, DrawOn.Series["MeanMarker"].LabelForeColor);
            DrawOn.Series["Trendline"].Color = Color.FromArgb(0, DrawOn.Series["Trendline"].Color);
            DrawOn.Series["Trendline"].LabelBackColor = Color.FromArgb(0, DrawOn.Series["Trendline"].LabelBackColor);
            DrawOn.Series["Trendline"].LabelForeColor = Color.FromArgb(0, DrawOn.Series["Trendline"].LabelForeColor);
            Series s25x = DrawOn.Series.FindByName("Percentile25_X");
            s25x.Color = Color.FromArgb(0, s25x.Color);
            Series s75x = DrawOn.Series.FindByName("Percentile75_X");
            s75x.Color = Color.FromArgb(0, s75x.Color);
            Series s25y = DrawOn.Series.FindByName("Percentile25_Y");
            s25y.Color = Color.FromArgb(0, s25y.Color);
            Series s75y = DrawOn.Series.FindByName("Percentile75_Y");
            s75y.Color = Color.FromArgb(0, s75y.Color);
            Series sMeanx = DrawOn.Series.FindByName("PercentileMean_X");
            sMeanx.Color = Color.FromArgb(0, sMeanx.Color);
            Series sMeany = DrawOn.Series.FindByName("PercentileMean_Y");
            sMeany.Color = Color.FromArgb(0, sMeany.Color);

            DrawSeries.Color = Color.FromArgb(255, DrawSeries.Color);
            DrawOn.Series["CorrelSeries"].Label = "";
        }

        public void ResetChartFormat()
        {
            //Reset the background color
            DrawOn.SendToBack();
            DrawArea.BackColor = ExistingCanvasColor;
            //Reset the marker colors
            DrawOn.Series["CorrelSeries"].Color = Color.FromArgb(255, DrawOn.Series["CorrelSeries"].Color);
            DrawOn.Series["CorrelSeries"].LabelBackColor = Color.FromArgb(255, DrawOn.Series["CorrelSeries"].LabelBackColor);
            DrawOn.Series["CorrelSeries"].LabelForeColor = Color.FromArgb(255, DrawOn.Series["CorrelSeries"].LabelForeColor);
            DrawOn.Series["MeanMarker"].Color = Color.FromArgb(255, DrawOn.Series["MeanMarker"].Color);
            DrawOn.Series["MeanMarker"].LabelBackColor = Color.FromArgb(255, DrawOn.Series["MeanMarker"].LabelBackColor);
            DrawOn.Series["MeanMarker"].LabelForeColor = Color.FromArgb(255, DrawOn.Series["MeanMarker"].LabelForeColor);
            DrawOn.Series["Trendline"].Color = Color.FromArgb(255, DrawOn.Series["Trendline"].Color);
            DrawOn.Series["Trendline"].LabelBackColor = Color.FromArgb(255, DrawOn.Series["Trendline"].LabelBackColor);
            DrawOn.Series["Trendline"].LabelForeColor = Color.FromArgb(255, DrawOn.Series["Trendline"].LabelForeColor);

            DrawSeries.Color = Color.FromArgb(0, DrawSeries.Color);
            Series s25x = DrawOn.Series.FindByName("Percentile25_X");
            s25x.Color = Color.FromArgb(100, s25x.Color);
            Series s75x = DrawOn.Series.FindByName("Percentile75_X");
            s75x.Color = Color.FromArgb(100, s75x.Color);
            Series s25y = DrawOn.Series.FindByName("Percentile25_Y");
            s25y.Color = Color.FromArgb(100, s25y.Color);
            Series s75y = DrawOn.Series.FindByName("Percentile75_Y");
            s75y.Color = Color.FromArgb(100, s75y.Color);
            Series sMeanx = DrawOn.Series.FindByName("PercentileMean_X");
            sMeanx.Color = Color.FromArgb(255, sMeanx.Color);
            Series sMeany = DrawOn.Series.FindByName("PercentileMean_Y");
            sMeany.Color = Color.FromArgb(255, sMeany.Color);

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
            Random rando = new Random();
            for (int i = 0; i < 3; i++)
            {
                //Drop 3x points at once
                int x = System.Windows.Forms.Cursor.Position.X;
                int y = System.Windows.Forms.Cursor.Position.Y;
                //Make the standard deviation for the offset based on the axis chart stdev?

                CorrelationForm chartParent = (CorrelationForm)DrawOn.Parent;
                Chart xChart = (Chart)chartParent.Controls.Find("xAxisChart", false).First();
                Chart yChart = (Chart)chartParent.Controls.Find("yAxisChart", false).First();
                //Pixel offsets for the "paint can" look
                Accord.Statistics.Distributions.Univariate.NormalDistribution offset_x_Dist = new Accord.Statistics.Distributions.Univariate.NormalDistribution(0, 25);
                Accord.Statistics.Distributions.Univariate.NormalDistribution offset_y_Dist = new Accord.Statistics.Distributions.Univariate.NormalDistribution(0, 25);
                
                int x_offset = Convert.ToInt32(offset_x_Dist.InverseDistributionFunction(rando.NextDouble()));
                int y_offset = Convert.ToInt32(offset_y_Dist.InverseDistributionFunction(rando.NextDouble()));
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
        }

        private double GetPaintTimerInterval(DataPoint dp)
        {
            //Get the height of the x-axis chart at this point.
            CorrelationForm chartParent = (CorrelationForm)DrawOn.Parent;
            Chart xChart = (Chart)chartParent.Controls.Find("xAxisChart", false).First();
            Chart yChart = (Chart)chartParent.Controls.Find("yAxisChart", false).First();
            double xChart_Value = chartParent.CorrelDist1.GetPDF_Value(dp.XValue) / 2;
            double xChart_Max = chartParent.CorrelDist1.GetPDF_MaxHeight();
            double xChart_Ratio = xChart_Value / xChart_Max;
            
            double yChart_Value = chartParent.CorrelDist2.GetPDF_Value(dp.YValues.First()) / 2;
            double yChart_Max = chartParent.CorrelDist2.GetPDF_MaxHeight();
            double yChart_Ratio = yChart_Value / yChart_Max;

            double returnValue = (1 - (xChart_Ratio + yChart_Ratio)) * PaintTimerDefaultRate;
            return returnValue;     //A lower number is faster because it is the number of milliseconds between point drops
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
                PaintTimer.Interval = GetPaintTimerInterval(newDataPoint);
            }

            ////PLAN: NEED TO CONVERT the Point's coords to axis values (axisToValue inside paint event?), create a DataPoint off the derived x & y
            //Check if you are within the InnerPlot area - if so, add and return true. If not, return false.
        }

        
    }
}
