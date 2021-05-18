using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Accord.Statistics.Distributions.Univariate;
using System.Windows.Forms.DataVisualization.Charting;
using Excel = Microsoft.Office.Interop.Excel;


namespace CorrelationTest
{
    public partial class CorrelationForm : Form
    {
        public enum CoefficientBox_ErrorType
        {
            Feasibility,
            Transitivity,
            PSD,
            Conformal
        }
        private bool RefreshBreak { get; set; } = false;
        private Chart yAxisChart {get;set;}
        private Tuple<Point, Point>[] hoverPoints { get; set; } = new Tuple<Point, Point>[10];
        private Label HoverLabel_H1 { get; set; }
        private Label HoverLabel_H2 { get; set; }
        private Label HoverLabel_H3 { get; set; }
        private Label HoverLabel_H4 { get; set; }
        private Label HoverLabel_H5 { get; set; }
        private Label HoverLabel_V1 { get; set; }
        private Label HoverLabel_V2 { get; set; }
        private Label HoverLabel_V3 { get; set; }
        private Label HoverLabel_V4 { get; set; }
        private Label HoverLabel_V5 { get; set; }
        Dictionary<string, dynamic> Spacing { get; set; } = new Dictionary<string, dynamic>();
        Dictionary<string, DataPoint> PercentilePoints { get; set; } = new Dictionary<string, DataPoint>();
        private CoefficientBox_ErrorType maxConstraint { get; set; }
        private CoefficientBox_ErrorType minConstraint { get; set; }
        private Tuple<double, double> trans_bounds { get; set; }
        private Tuple<double, double> feasibility_bounds { get; set; }
        private List<DataPoint> CorrelScatterPoints { get; set; } = new List<DataPoint>();
        private Color existingColor { get; set; }
        private Color existingColor_Markers { get; set; }
        private int helperStage { get; set; }
        private TextBox textboxMinimum = new TextBox();
        private TextBox textboxMidpoint = new TextBox();
        private TextBox textboxMaximum = new TextBox();
        private DrawnCorrelation DrawnCorrel { get; set; }
        private DrawingTool DrawTool { get; set; }
        private bool DrawingMode = false;
        private bool UpDownEnabled { get; set; }
        private Label labelHelper = new Label();
        private Label label_coefErrors { get; set; }
        private decimal lastValue { get; set; }
        private bool errorState_CoefficientBox { get; set; }
        private IEstimateDistribution CorrelDist1 { get; set; }
        private IEstimateDistribution CorrelDist2 { get; set; }

        public CorrelationForm(IEstimateDistribution correlDist1, IEstimateDistribution correlDist2)
        {
            this.CorrelDist1 = correlDist1;
            this.CorrelDist2 = correlDist2;

        

            InitializeComponent();
        }

        private void ReloadCorrelScatter()
        {
            //Clears the drawing by reloading the default CorrelScatter again
            this.CorrelScatter = new System.Windows.Forms.DataVisualization.Charting.Chart();
            this.CorrelScatter.Series[0].MarkerStyle = MarkerStyle.Circle;
            //Set the axis scale
            this.CorrelScatter.ChartAreas[0].AxisX.LabelStyle.Format = "0.0";
            this.CorrelScatter.ChartAreas[0].AxisY.LabelStyle.Format = "0.0";
            this.CorrelScatter.ChartAreas[0].AxisY.Enabled = AxisEnabled.False;
            this.CorrelScatter.ChartAreas[0].AxisY.Enabled = AxisEnabled.True;
            //this.CorrelScatter.ChartAreas[0].AxisX.Interval = .5;
            this.CorrelScatter.ChartAreas[0].AxisX.Minimum = CorrelDist1.GetMinimum();
            this.CorrelScatter.ChartAreas[0].AxisX.Maximum = CorrelDist1.GetMaximum();
            //this.CorrelScatter.ChartAreas[0].AxisY2.Interval = .5;
            this.CorrelScatter.ChartAreas[0].AxisY.Minimum = CorrelDist2.GetMinimum();
            this.CorrelScatter.ChartAreas[0].AxisY.Maximum = CorrelDist2.GetMaximum();


            foreach (DataPoint dp in CorrelScatterPoints)
            {
                this.CorrelScatter.Series["CorrelSeries"].Points.AddXY(dp.XValue, dp.YValues[0]);
            }
        }

        private void CorrelationForm_Load(object sender, EventArgs e)
        {
            Sheets.CorrelationSheet CorrelSheet = Sheets.CorrelationSheet.ConstructFromXlCorrelationSheet();
            CorrelScatter.Height = 750;
            CorrelScatter.Width = 750;
            CorrelScatter.ChartAreas[0].Position = new ElementPosition(5, 3, 90, 90);
            CorrelScatter.ChartAreas[0].InnerPlotPosition = new ElementPosition(5, 3, 90, 90);
            //Create & set example points
            Random rando = new Random();
            for (int i = 1; i < 500; i++)
            {
                //double input = ((double)i) / 100;
                double x = CorrelDist1.GetInverse(rando.NextDouble());
                double y = CorrelDist2.GetInverse(rando.NextDouble());
                CorrelScatterPoints.Add(new DataPoint(x, y));
                this.CorrelScatter.Series["CorrelSeries"].Points.AddXY(x, y);
            }

            this.CorrelScatter.Series["CorrelSeries"].MarkerStyle = MarkerStyle.Circle;
            this.CorrelScatter.Series["CorrelSeries"].IsVisibleInLegend = false;
            //Set the axis scale
            this.CorrelScatter.ChartAreas[0].AxisX.LabelStyle.Format = "0.0";
            this.CorrelScatter.ChartAreas[0].AxisY2.LabelStyle.Format = "0.0";
            
            //this.CorrelScatter.ChartAreas[0].AxisX.Interval = .5;
            this.CorrelScatter.ChartAreas[0].AxisX.Minimum = CorrelDist1.GetMinimum();
            this.CorrelScatter.ChartAreas[0].AxisX.Maximum = CorrelDist1.GetMaximum();
            //this.CorrelScatter.ChartAreas[0].AxisY2.Interval = .5;
            this.CorrelScatter.ChartAreas[0].AxisY.Minimum = CorrelDist2.GetMinimum();
            this.CorrelScatter.ChartAreas[0].AxisY.Maximum = CorrelDist2.GetMaximum();

            this.CorrelScatter.ChartAreas[0].AxisY2.Minimum = CorrelDist2.GetMinimum();
            this.CorrelScatter.ChartAreas[0].AxisY2.Maximum = CorrelDist2.GetMaximum();
            //this.CorrelScatter.ChartAreas[0].AxisY2.MajorGrid.Interval = 

            //Mean marker series
            Series meanMarker = new Series();
            meanMarker.ChartType = SeriesChartType.Point;
            meanMarker.Points.AddXY(CorrelDist1.GetMean(), CorrelDist2.GetMean());
            meanMarker.Color = Color.FromArgb(255, 0, 0, 0);
            meanMarker.MarkerStyle = MarkerStyle.Square;
            meanMarker.MarkerSize = 10;
            
            CorrelScatter.Series.Add(meanMarker);
            
            Excel.Range xlSelection = ThisAddIn.MyApp.Selection;
            int index1 = xlSelection.Row - (CorrelSheet.xlMatrixCell.Row + 1);
            int index2 = xlSelection.Column - CorrelSheet.xlMatrixCell.Column;


            trans_bounds = CorrelSheet.CorrelMatrix.GetTransitivityBounds(index1, index2);  //<min, max>
            feasibility_bounds = CorrelSheet.CorrelMatrix.GetFeasibilityBounds(CorrelDist1, CorrelDist2);
            numericUpDown_CorrelValue.TextAlign = HorizontalAlignment.Center;
            numericUpDown_CorrelValue.DecimalPlaces = 2;

            double bindingMin = Math.Max(feasibility_bounds.Item1, trans_bounds.Item1);
            if (bindingMin == -1)
                minConstraint = CoefficientBox_ErrorType.Conformal;
            else if (bindingMin == feasibility_bounds.Item1)
                minConstraint = CoefficientBox_ErrorType.Feasibility;
            else
                minConstraint = CoefficientBox_ErrorType.Transitivity;

            double bindingMax = Math.Min(feasibility_bounds.Item2, trans_bounds.Item2);
            if (bindingMax == 1)
                maxConstraint = CoefficientBox_ErrorType.Conformal;
            else if (bindingMax == feasibility_bounds.Item2)
                maxConstraint = CoefficientBox_ErrorType.Feasibility;
            else
                maxConstraint = CoefficientBox_ErrorType.Transitivity;

            //NumericUpDown numericUpDown_CorrelValue = new NumericUpDown();
            numericUpDown_CorrelValue.Minimum = Decimal.Ceiling((Convert.ToDecimal(bindingMin) * 100)) / 100;
            numericUpDown_CorrelValue.Maximum = Decimal.Floor((Convert.ToDecimal(bindingMax) * 100)) / 100;
            //groupBox_CorrelCoef.Controls.Add(numericUpDown_CorrelValue);

            numericUpDown_CorrelValue.Increment = Convert.ToDecimal(0.01);

            this.label_coefErrors = new Label();
            label_coefErrors.AutoSize = false;
            label_coefErrors.Width = this.groupBox_CorrelCoef.Width - 4;
            this.groupBox_CorrelCoef.Controls.Add(label_coefErrors);
            label_coefErrors.Top = numericUpDown_CorrelValue.Bottom;
            label_coefErrors.Height = this.groupBox_CorrelCoef.Height - label_coefErrors.Top - 2;
            label_coefErrors.Left = 2;
            label_coefErrors.Padding = new Padding(6);

            double existingValue = Convert.ToDouble(ThisAddIn.MyApp.Selection.value);
            if (existingValue >= trans_bounds.Item1 && existingValue <= trans_bounds.Item2)
            {
                numericUpDown_CorrelValue.Value = Convert.ToDecimal(existingValue); //Keep existing matrix value
                CoefficientBox_Reset();
                lastValue = numericUpDown_CorrelValue.Value;
            }
            else if (existingValue < trans_bounds.Item1 || existingValue < feasibility_bounds.Item1)
            {
                numericUpDown_CorrelValue.Value = Convert.ToDecimal(trans_bounds.Item1);  //Set to min
                if(existingValue <= -1)
                    CoefficientBox_FlagError(CoefficientBox_ErrorType.Conformal);
                else if(existingValue <= feasibility_bounds.Item1)
                    CoefficientBox_FlagError(CoefficientBox_ErrorType.Feasibility);
                else if(existingValue <= trans_bounds.Item1)
                    CoefficientBox_FlagError(CoefficientBox_ErrorType.Transitivity);
                lastValue = numericUpDown_CorrelValue.Value;
            }
            else if (existingValue > trans_bounds.Item2 || existingValue > feasibility_bounds.Item2)
            {
                numericUpDown_CorrelValue.Value = Convert.ToDecimal(trans_bounds.Item2);  //Set to max
                if (existingValue >= 1)
                    CoefficientBox_FlagError(CoefficientBox_ErrorType.Conformal);
                else if(existingValue >= feasibility_bounds.Item2)
                    CoefficientBox_FlagError(CoefficientBox_ErrorType.Feasibility);
                else if(existingValue >= trans_bounds.Item2)
                    CoefficientBox_FlagError(CoefficientBox_ErrorType.Transitivity);
                lastValue = numericUpDown_CorrelValue.Value;
            }
            else
            {
                throw new Exception("Unhandled initial condition");
            }

            LoadXAxisDistribution();
            LoadYAxisDistribution();
            SetupHelper();
            SetupHoverPoints();

        }

        private void LoadYAxisDistribution()
        {
            //Build a series off the distribution
            System.IO.MemoryStream myStream = new System.IO.MemoryStream();
            this.yAxisChart = new Chart();
            
            xAxisChart.Serializer.Save(myStream);
            yAxisChart.Serializer.Load(myStream);

            yAxisChart.Series.Clear();
            Series Series1 = new Series();
            yAxisChart.Series.Add(Series1);
            yAxisChart.Series["Series1"].ChartType = SeriesChartType.Bar;
            yAxisChart.Width = xAxisChart.Height;
            yAxisChart.Left = CorrelScatter.Left - yAxisChart.Width;
            
            yAxisChart.Top = CorrelScatter.Top;
            yAxisChart.Height = CorrelScatter.Height;

            //yAxisChart.ChartAreas[0].Position.X = 0;
            yAxisChart.ChartAreas[0].Position = new ElementPosition(5, 3, 90, 90);
            yAxisChart.ChartAreas[0].InnerPlotPosition = new ElementPosition(5, 3, 90, 90);
            //yAxisChart.ChartAreas[0].InnerPlotPosition.X = 0;
            //yAxisChart.ChartAreas[0].InnerPlotPosition.Width = 100; //xAxisChart.ChartAreas[0].InnerPlotPosition.Height;

            //yAxisChart.ChartAreas[0].Position.Y = CorrelScatter.ChartAreas[0].Position.Y;
            //yAxisChart.ChartAreas[0].InnerPlotPosition.Y = CorrelScatter.ChartAreas[0].InnerPlotPosition.Y;
            //yAxisChart.ChartAreas[0].InnerPlotPosition.Height = CorrelScatter.ChartAreas[0].InnerPlotPosition.Height;

            yAxisChart.Series["Series1"].YValuesPerPoint = 1;
            //yAxisChart.ChartAreas[0].AxisX.Interval = 0.5;
            yAxisChart.Series["Series1"].IsVisibleInLegend = false;
            yAxisChart.Series["Series1"]["PixelPointWidth"] = "3";

            int steps = 400;
            double minimum = CorrelDist2.GetMinimum();
            double maximum = CorrelDist2.GetMaximum();
            double step = (maximum - minimum) / steps;

            yAxisChart.ChartAreas[0].AxisX.Minimum = CorrelScatter.ChartAreas[0].AxisY2.Minimum;
            yAxisChart.ChartAreas[0].AxisX.Maximum = CorrelScatter.ChartAreas[0].AxisY2.Maximum;

            for (int i = 0; i < steps; i++)
            {                
                double x = minimum + step * i;
                double y = CorrelDist2.GetPDF_Value(x);
                yAxisChart.Series["Series1"].Points.AddXY(x, y);
            }

            double meanValue = CorrelDist2.GetMean();
            //Find the point in the series that is closest to the mean.
            var meanDistances = from DataPoint dp in yAxisChart.Series["Series1"].Points select new Tuple<DataPoint, double>(dp, Math.Abs(dp.XValue - meanValue));
            DataPoint meanPoint = meanDistances.OrderBy(t => t.Item2).First().Item1;
            PercentilePoints.Add("Y_MeanPoint", meanPoint);
            meanPoint.Color = Color.FromArgb(0, 0, 0);

            double lowValue = CorrelDist2.GetInverse(0.25);
            var lowDistances = from DataPoint dp in yAxisChart.Series["Series1"].Points select new Tuple<DataPoint, double>(dp, Math.Abs(dp.XValue - lowValue));
            DataPoint lowPoint = lowDistances.OrderBy(t => t.Item2).First().Item1;
            PercentilePoints.Add("Y_LowPoint", lowPoint);
            lowPoint.Color = Color.FromArgb(50, 50, 50);

            double highValue = CorrelDist2.GetInverse(0.75);
            var highDistances = from DataPoint dp in yAxisChart.Series["Series1"].Points select new Tuple<DataPoint, double>(dp, Math.Abs(dp.XValue - highValue));
            DataPoint highPoint = highDistances.OrderBy(t => t.Item2).First().Item1;
            PercentilePoints.Add("Y_HighPoint", highPoint);
            highPoint.Color = Color.FromArgb(50, 50, 50);

            //Find the point in the series that is closest to the mean.
            var distances = from DataPoint dp in yAxisChart.Series["Series1"].Points select new Tuple<DataPoint, double>(dp, Math.Abs(dp.XValue - meanValue));
            var ordered = distances.OrderBy(t => t.Item2);
            DataPoint closestPoint = ordered.First().Item1;
            closestPoint.Color = Color.FromArgb(0, 0, 0);
            closestPoint.BackSecondaryColor = Color.FromArgb(0, 0, 0);

            yAxisChart.ChartAreas[0].AxisX.Interval = CorrelScatter.ChartAreas[0].AxisY.Interval;

            yAxisChart.ChartAreas[0].AxisX.LabelStyle.Format = "0.0";
            yAxisChart.ChartAreas[0].AxisX.LabelStyle.Enabled = false;
            yAxisChart.ChartAreas[0].AxisY.LabelStyle.Format = "0.0";
            yAxisChart.ChartAreas[0].AxisY.IsReversed = true;

            yAxisChart.Paint += YAxisChart_Paint;

            this.Controls.Add(yAxisChart);
        }

        private void LoadXAxisDistribution()
        {
            //Build a series off the distribution
            this.xAxisChart.Left = CorrelScatter.Left;
            this.xAxisChart.Top = CorrelScatter.Top - 150;
            this.xAxisChart.Height = 150;
            this.xAxisChart.Width = CorrelScatter.Width;

            xAxisChart.ChartAreas[0].Position = new ElementPosition(5, 3, 90, 90);
            xAxisChart.ChartAreas[0].InnerPlotPosition = new ElementPosition(5, 3, 90, 90);

            this.xAxisChart.Series["Series1"].YValuesPerPoint = 1;
            xAxisChart.Series["Series1"]["PixelPointWidth"] = "3";

            int steps = 400;
            double minimum = CorrelDist1.GetMinimum();
            double maximum = CorrelDist1.GetMaximum();
            double step = (maximum - minimum) / steps;

            this.xAxisChart.ChartAreas[0].AxisX.Minimum = minimum;
            this.xAxisChart.ChartAreas[0].AxisX.Maximum = maximum;

            for (int i = 0; i < steps; i++)
            {
                double x = minimum + step * i;
                double y = CorrelDist1.GetPDF_Value(x);
                xAxisChart.Series["Series1"].Points.AddXY(x, y);
            }
            double meanValue = CorrelDist1.GetMean();
            //Find the point in the series that is closest to the mean.
            var meanDistances = from DataPoint dp in xAxisChart.Series["Series1"].Points select new Tuple<DataPoint, double>(dp, Math.Abs(dp.XValue - meanValue));
            DataPoint meanPoint = meanDistances.OrderBy(t => t.Item2).First().Item1;
            PercentilePoints.Add("X_MeanPoint", meanPoint);
            meanPoint.Color = Color.FromArgb(0, 0, 0);

            double lowValue = CorrelDist1.GetInverse(0.25);
            var lowDistances = from DataPoint dp in xAxisChart.Series["Series1"].Points select new Tuple<DataPoint, double>(dp, Math.Abs(dp.XValue - lowValue));
            DataPoint lowPoint = lowDistances.OrderBy(t => t.Item2).First().Item1;
            PercentilePoints.Add("X_LowPoint", lowPoint);
            lowPoint.Color = Color.FromArgb(50, 50, 50);

            double highValue = CorrelDist1.GetInverse(0.75);
            var highDistances = from DataPoint dp in xAxisChart.Series["Series1"].Points select new Tuple<DataPoint, double>(dp, Math.Abs(dp.XValue - highValue));
            DataPoint highPoint = highDistances.OrderBy(t => t.Item2).First().Item1;
            PercentilePoints.Add("X_HighPoint", highPoint);
            highPoint.Color = Color.FromArgb(50, 50, 50);
            

            xAxisChart.ChartAreas[0].AxisY.IntervalAutoMode = IntervalAutoMode.VariableCount;
            this.xAxisChart.ChartAreas[0].AxisX.Interval = CorrelScatter.ChartAreas[0].AxisX.Interval;
            xAxisChart.ChartAreas[0].AxisX.LabelStyle.Format = "0.0";
            xAxisChart.ChartAreas[0].AxisY.LabelStyle.Format = "0.0";
            xAxisChart.ChartAreas[0].AxisX.LabelStyle.Enabled = false;

            //Load the percentile lines
            xAxisChart.Paint += XAxisChart_Paint;
        }

        private enum Percentile
        {
            Mean,
            Low,
            Mid,
            High
        }

        private void DrawPercentile(Percentile percentile, QuintantOrientation orientation)
        {
            if(orientation == QuintantOrientation.Horizontal)
            {
                switch (percentile)
                {
                    case Percentile.Mean:
                        break;
                    case Percentile.Low:
                        break;
                    case Percentile.Mid:
                        break;
                    case Percentile.High:
                        break;
                    default:
                        throw new Exception("Unexpected percentile");
                }
            }
            else if(orientation == QuintantOrientation.Vertical)
            {
                switch (percentile)
                {
                    case Percentile.Mean:
                        break;
                    case Percentile.Low:
                        break;
                    case Percentile.Mid:
                        break;
                    case Percentile.High:
                        break;
                    default:
                        throw new Exception("Unexpected percentile");
                }
            }
            else
            {
                throw new Exception("Unexpected orientation");
            }
        }

        public void CoefficientBox_Reset()
        {
            //DEFAULT "Info" yellow
            groupBox_CorrelCoef.BackColor = Color.FromArgb(255, 255, 225);
            this.label_coefErrors.Text = "";
            errorState_CoefficientBox = false;
        }

        public void CoefficientBox_FlagError(CoefficientBox_ErrorType errorType)
        {
            //ERROR red
            errorState_CoefficientBox = true;
            groupBox_CorrelCoef.BackColor = Color.FromArgb(255, 124, 128);

            switch (errorType)
            {
                case CoefficientBox_ErrorType.Conformal:
                    this.label_coefErrors.Text = "More extreme values violate conformality";
                    break;
                case CoefficientBox_ErrorType.Transitivity:
                    this.label_coefErrors.Text = "More extreme values violate transitivity";
                    break;
                case CoefficientBox_ErrorType.Feasibility:
                    this.label_coefErrors.Text = "More extreme values violate feasibility";
                    break;
                case CoefficientBox_ErrorType.PSD:
                    this.label_coefErrors.Text = "Value makes matrix fail PSD check";
                    break;
                default:
                    throw new Exception("Unknown error type");
            }
        }
        
        private void btn_Cancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btn_saveClose_Click(object sender, EventArgs e)
        {
            //Need to identify & replace the number in the matrix.
            //The number itself should have the bounds of what it can be set to established on visual launch.
            //Pull the correl sheet object for visualization again
            Sheets.CorrelationSheet correlSheet = Sheets.CorrelationSheet.ConstructFromXlCorrelationSheet();
            Excel.Range xlSelection = ThisAddIn.MyApp.Selection;
            int index1 = xlSelection.Row - (correlSheet.xlMatrixCell.Row+1);
            int index2 = xlSelection.Column - correlSheet.xlMatrixCell.Column;
            int temp;
            if (index1 > index2)    //If the selection is in the lower triangular, edit the value in the upper triangular instead
            {
                temp = index2;
                index2 = index1;
                index1 = temp;
            }
            double newValue = Convert.ToDouble(numericUpDown_CorrelValue.Value);
            correlSheet.CorrelMatrix.SetCorrelation(index1, index2, newValue);
            if(!Sheets.CorrelationSheet.CheckMatrixForErrors(correlSheet))
                correlSheet.CorrelMatrix.PrintToSheet(correlSheet.xlMatrixCell);
            else
            {
                throw new Exception("Matrix errors");
            }
            this.Close();
        }

        private void groupBox_CorrelCoef_Enter(object sender, EventArgs e)
        {

        }

        private void numericUpDown_CorrelValue_MouseDown(object sender, MouseEventArgs e)
        {
            //Instead of doing this big comparison song and dance, I should just bool save what the constraints are...
            Decimal.TryParse(Convert.ToString(numericUpDown_CorrelValue.Value), out decimal currentValue);
            if (lastValue == currentValue)
            {
                if (currentValue >= numericUpDown_CorrelValue.Maximum)
                {
                    CoefficientBox_FlagError(maxConstraint);
                }
                else if (numericUpDown_CorrelValue.Value <= numericUpDown_CorrelValue.Minimum)
                {
                    CoefficientBox_FlagError(minConstraint);
                }
            }
            else
            {
                CoefficientBox_Reset();
            }
            lastValue = numericUpDown_CorrelValue.Value;
        }

        private void numericUpDown_CorrelValue_ValueChanged(object sender, EventArgs e)
        {
            //Save the old value
            //CoefficientBox_Reset();
        }

        private void SetupHelper()
        {
            int vertical = CorrelScatter.Height / 2;
            int min = CorrelScatter.Width / 5;
            int mid = min * 2;
            int max = min * 3;

            textboxMinimum.Top = vertical;
            textboxMinimum.Left = CorrelScatter.Left + min;
            textboxMinimum.Height = 50;
            textboxMinimum.Width = 50;
            textboxMidpoint.Top = vertical;
            textboxMidpoint.Left = CorrelScatter.Left + mid + 30;
            textboxMidpoint.Height = 50;
            textboxMidpoint.Width = 50;
            textboxMaximum.Top = vertical;
            textboxMaximum.Left = CorrelScatter.Left + max + 60;
            textboxMaximum.Height = 50;
            textboxMaximum.Width = 50;

        }

        private void btn_LaunchHelper_Click(object sender, EventArgs e)
        {
            if(helperStage == 0)
            {
                //Dis-enable the other controls
                this.btn_LaunchDrawCorrelation.Enabled = false;
                this.btn_saveClose.Enabled = false;
                this.UpDownEnabled = this.numericUpDown_CorrelValue.Enabled;
                if (this.UpDownEnabled)
                    this.numericUpDown_CorrelValue.Enabled = false;
                //Load minimum
                existingColor = CorrelScatter.ChartAreas[0].BackColor;
                CorrelScatter.ChartAreas[0].BackColor = Color.FromArgb(195, 195, 195);
                this.btn_LaunchHelper.Text = ">> Next >>";
                this.Controls.Add(textboxMinimum);
                textboxMinimum.BringToFront();
                labelHelper.AutoSize = true;
                labelHelper.Top = textboxMinimum.Top - 50;
                labelHelper.Left = textboxMinimum.Left;
                labelHelper.Text = $"If X is {CorrelDist1.GetMinimum()}, what do you expect Y to be?";
                this.Controls.Add(labelHelper);
                labelHelper.BringToFront();
                helperStage++;
            }
            else if(helperStage == 1)
            {
                textboxMinimum.Enabled = false;
                Color existingColor = CorrelScatter.ChartAreas[0].BackColor;
                this.Controls.Add(textboxMidpoint);
                textboxMidpoint.BringToFront();
                labelHelper.Top = textboxMidpoint.Top - 50;
                labelHelper.Left = textboxMidpoint.Left;
                labelHelper.Text = $"If X is {(CorrelDist1.GetMaximum() - CorrelDist1.GetMinimum()) / 2}, what do you expect Y to be?";
                labelHelper.BringToFront();
                helperStage++;
            }
            else if (helperStage == 2)
            {
                textboxMidpoint.Enabled = false;
                Color existingColor = CorrelScatter.ChartAreas[0].BackColor;
                this.Controls.Add(textboxMaximum);
                textboxMaximum.BringToFront();
                labelHelper.Top = textboxMaximum.Top - 50;
                labelHelper.Left = textboxMaximum.Left;
                labelHelper.Text = $"If X is {CorrelDist1.GetMaximum()}, what do you expect Y to be?";
                labelHelper.BringToFront();
                helperStage++;
            }
            else if (helperStage == 3)
            {
                //Save the values
                bool t1 = Double.TryParse(textboxMinimum.Text, out double minVal);
                bool t2 = Double.TryParse(textboxMidpoint.Text, out double midVal);
                bool t3 = Double.TryParse(textboxMaximum.Text, out double maxVal);
                
                if(t1&&t2&&t3)
                {
                    //all three contain convertible values
                    
                    //COMPUTE THE CORRELATION HERE

                }
                //Remove the textboxes
                this.Controls.Remove(textboxMinimum);
                this.Controls.Remove(textboxMidpoint);
                this.Controls.Remove(textboxMaximum);
                this.Controls.Remove(labelHelper);
                this.btn_LaunchHelper.Text = "Use Guided Correlation";
                //Return the color to normal
                CorrelScatter.ChartAreas[0].BackColor = existingColor;
                //Compute the line?
                //But the slope != the correlation...
                //So what am I doing here?

                this.btn_LaunchDrawCorrelation.Enabled = true;
                this.btn_saveClose.Enabled = true;
                this.numericUpDown_CorrelValue.Enabled = UpDownEnabled;

                
                helperStage = 0;
            }
        }

        private void btn_LaunchDrawCorrelation_Click(object sender, EventArgs e)
        {
            if (!DrawingMode)
            {
                //Turn on DrawingMode
                DrawingMode = true;
                ChartArea correlScatterArea = CorrelScatter.ChartAreas[0];
                this.DrawTool = new DrawingTool(ref CorrelScatter, ref correlScatterArea);
                DrawTool.FormatChartForDrawing();
                foreach(Label qLab in CorrelScatter.Controls)
                {
                    qLab.Hide();
                }
                //this.DrawnCorrel = new DrawnCorrelation();
                //Disable the other buttons
                this.btn_LaunchHelper.Enabled = false;
                this.btn_saveClose.Enabled = false;
                this.UpDownEnabled = this.numericUpDown_CorrelValue.Enabled;
                if(this.UpDownEnabled)
                    this.numericUpDown_CorrelValue.Enabled = false;
                btn_LaunchDrawCorrelation.Text = "Done Drawing";
            }
            else
            {
                //Turn off DrawingMode
                DrawTool.ResetChartFormat();
                CorrelScatter.Series.Remove(DrawTool.DrawSeries);
                foreach (Label qLab in CorrelScatter.Controls)
                {
                    qLab.Show();
                }
                this.DrawTool = null;                
                this.btn_LaunchHelper.Enabled = true;
                this.btn_saveClose.Enabled = true;
                this.numericUpDown_CorrelValue.Enabled = this.UpDownEnabled;    //Reset to original state
                btn_LaunchDrawCorrelation.Text = "Draw Correlation";
                DrawingMode = false;
            }
        }

        private void CorrelScatter_MouseClick(object sender, MouseEventArgs e)
        {
            if (DrawingMode)
            {
                DrawTool.AddPoint(e.Location);
            }
            
        }

        private void CorrelScatter_Paint(object sender, PaintEventArgs e)
        {
            if (DrawingMode)
            {
                DrawTool.GetXAxisMinMax();       //Does this work? Called from paint event, but indirectly...
                DrawTool.GetYAxisMinMax();       //Does this work? Called from paint event, but indirectly...
                
                if (DrawTool.DrawSeries.Points.Count() > 1)
                {
                    //PLAN: Create a new series in CorrelScatter, use the drawing tool to add to it, then refresh here
                    //DrawTool.Refresh();
                }
                else if(DrawTool == null)
                {
                    this.ReloadCorrelScatter();
                }
            }
            else
            {
                if (Spacing.ContainsKey("X_MeanPoint_Abs"))
                {
                    double x_mean = Spacing["X_MeanPoint_Abs"];
                    Point[] points = new Point[2];
                    points[0] = new Point(Convert.ToInt32(x_mean), Spacing["chartInnerPlot_Abs_Top"]);
                    points[1] = new Point(Convert.ToInt32(x_mean), Spacing["chartInnerPlot_Abs_Bottom"]);
                    Pen pen_mean = new Pen(Color.FromArgb(150, 0, 0, 0));
                    pen_mean.DashStyle = System.Drawing.Drawing2D.DashStyle.Dash;
                    pen_mean.Width = 3;
                    e.Graphics.DrawLine(pen_mean, points[0], points[1]);
                    //CorrelScatter.SendToBack();
                }
                if (Spacing.ContainsKey("X_LowPoint_Abs"))
                {
                    double x_low = Spacing["X_LowPoint_Abs"];
                    Point[] points = new Point[2];
                    points[0] = new Point(Convert.ToInt32(x_low), Spacing["chartInnerPlot_Abs_Top"]);
                    points[1] = new Point(Convert.ToInt32(x_low), Spacing["chartInnerPlot_Abs_Bottom"]);
                    Pen pen_low = new Pen(Color.FromArgb(150, 255, 99, 71));
                    pen_low.Width = 3;
                    e.Graphics.DrawLine(pen_low, points[0], points[1]);
                    //CorrelScatter.SendToBack();
                }
                if (Spacing.ContainsKey("X_HighPoint_Abs"))
                {
                    double x_high = Spacing["X_HighPoint_Abs"];
                    Point[] points = new Point[2];
                    points[0] = new Point(Convert.ToInt32(x_high), Spacing["chartInnerPlot_Abs_Top"]);
                    points[1] = new Point(Convert.ToInt32(x_high), Spacing["chartInnerPlot_Abs_Bottom"]);
                    Pen pen_high = new Pen(Color.FromArgb(150, 255, 99, 71));
                    pen_high.Width = 3;
                    e.Graphics.DrawLine(pen_high, points[0], points[1]);
                    //CorrelScatter.SendToBack();
                }
                if (Spacing.ContainsKey("Y_MeanPoint_Abs"))
                {
                    double y_mean = Spacing["Y_MeanPoint_Abs"];
                    Point[] points = new Point[2];
                    points[0] = new Point(Spacing["chartInnerPlot_Abs_Left"], Convert.ToInt32(y_mean));
                    points[1] = new Point(Spacing["chartInnerPlot_Abs_Right"], Convert.ToInt32(y_mean));
                    Pen pen_mean = new Pen(Color.FromArgb(150, 0, 0, 0));
                    pen_mean.Width = 3;
                    pen_mean.DashStyle = System.Drawing.Drawing2D.DashStyle.Dash;
                    e.Graphics.DrawLine(pen_mean, points[0], points[1]);
                    //CorrelScatter.SendToBack();
                }
                if (Spacing.ContainsKey("Y_LowPoint_Abs"))
                {
                    double y_low = Spacing["Y_LowPoint_Abs"];
                    Point[] points = new Point[2];
                    points[0] = new Point(Spacing["chartInnerPlot_Abs_Left"], Convert.ToInt32(y_low));
                    points[1] = new Point(Spacing["chartInnerPlot_Abs_Right"], Convert.ToInt32(y_low));
                    Pen pen_low = new Pen(Color.FromArgb(150, 255, 99, 71));
                    pen_low.Width = 3;
                    e.Graphics.DrawLine(pen_low, points[0], points[1]);
                    //CorrelScatter.SendToBack();
                }
                if (Spacing.ContainsKey("Y_HighPoint_Abs"))
                {
                    double y_high = Spacing["Y_HighPoint_Abs"];
                    Point[] points = new Point[2];
                    points[0] = new Point(Spacing["chartInnerPlot_Abs_Left"], Convert.ToInt32(y_high));
                    points[1] = new Point(Spacing["chartInnerPlot_Abs_Right"], Convert.ToInt32(y_high));
                    Pen pen_high = new Pen(Color.FromArgb(150, 255, 99, 71));
                    pen_high.Width = 3;
                    e.Graphics.DrawLine(pen_high, points[0], points[1]);
                    //CorrelScatter.SendToBack();
                }
            }
        }

        private Label ConstructLabel(int labelNumber, QuintantOrientation orientation)
        {
            Label HoverLabel = new Label();
            HoverLabel.BackColor = Color.FromArgb(25, 125, 125, 125);
            HoverLabel.Text = "";

            if(orientation == QuintantOrientation.Horizontal)
            {
                HoverLabel.Height = 50;
                HoverLabel.Width = Spacing["chartInnerPlot_Abs_Width"] / 5;
                HoverLabel.Left = Spacing["chartInnerPlot_Abs_Left"] + HoverLabel.Width * (labelNumber-1);
                HoverLabel.Top = Spacing["chartInnerPlot_Abs_Bottom"] - HoverLabel.Height;
            }
            else if(orientation == QuintantOrientation.Vertical)
            {
                HoverLabel.Height = Spacing["chartInnerPlot_Abs_Height"] / 5;
                HoverLabel.Width = 50;
                HoverLabel.Left = Spacing["chartInnerPlot_Abs_Right"] - HoverLabel.Width;
                HoverLabel.Top = Spacing["chartInnerPlot_Abs_Top"] + HoverLabel.Height * (labelNumber-1);
                
            }
            else
            {
                throw new Exception("Unknown orientation");
            }

            CorrelScatter.Controls.Add(HoverLabel);
            HoverLabel.BringToFront();

            return HoverLabel;
        }
        private void ConstructLabels()
        {
            HoverLabel_H1 = ConstructLabel(1, QuintantOrientation.Horizontal);
            HoverLabel_H2 = ConstructLabel(2, QuintantOrientation.Horizontal);
            HoverLabel_H3 = ConstructLabel(3, QuintantOrientation.Horizontal);
            HoverLabel_H4 = ConstructLabel(4, QuintantOrientation.Horizontal);
            HoverLabel_H5 = ConstructLabel(5, QuintantOrientation.Horizontal);

            HoverLabel_V1 = ConstructLabel(1, QuintantOrientation.Vertical);
            HoverLabel_V2 = ConstructLabel(2, QuintantOrientation.Vertical);
            HoverLabel_V3 = ConstructLabel(3, QuintantOrientation.Vertical);
            HoverLabel_V4 = ConstructLabel(4, QuintantOrientation.Vertical);
            HoverLabel_V5 = ConstructLabel(5, QuintantOrientation.Vertical);
        }

        private void SetupHoverPoints()
        {
            //Use transparent labels that appear when you hover over a hoverPoint.
            //Hovering over any given hoverPoint puts a border around that point
            
            Spacing.Add("chart_Abs_Width", CorrelScatter.Width);

            Spacing.Add("chartArea_Abs_Width", Convert.ToInt32(Spacing["chart_Abs_Width"] * (CorrelScatter.ChartAreas[0].Position.Width / 100)));
            Spacing.Add("chartArea_Rel_Left", CorrelScatter.ChartAreas[0].Position.X);
            Spacing.Add("chartArea_Abs_Left", Convert.ToInt32(Spacing["chart_Abs_Width"] * (Spacing["chartArea_Rel_Left"] / 100)));
            Spacing.Add("chartArea_Rel_Right", CorrelScatter.ChartAreas[0].Position.Right);
            Spacing.Add("chartArea_Abs_Right", Convert.ToInt32(Spacing["chart_Abs_Width"] * (Spacing["chartArea_Rel_Right"] / 100)));

            Spacing.Add("chartInnerPlot_Abs_Width", Convert.ToInt32(Spacing["chartArea_Abs_Width"] * (CorrelScatter.ChartAreas[0].InnerPlotPosition.Width / 100)));
            Spacing.Add("chartInnerPlot_Rel_Left", CorrelScatter.ChartAreas[0].InnerPlotPosition.X);
            Spacing.Add("chartInnerPlot_Abs_Left", Convert.ToInt32(Spacing["chartArea_Abs_Width"] * (Spacing["chartInnerPlot_Rel_Left"] / 100)) + Spacing["chartArea_Abs_Left"]);
            Spacing.Add("chartInnerPlot_Rel_Right", CorrelScatter.ChartAreas[0].InnerPlotPosition.Right);
            Spacing.Add("chartInnerPlot_Abs_Right", Convert.ToInt32(Spacing["chartArea_Abs_Width"] * (Spacing["chartInnerPlot_Rel_Right"] / 100)) + Spacing["chartArea_Abs_Left"]);


            Spacing.Add("chart_Abs_Height", CorrelScatter.Height);

            Spacing.Add("chartArea_Abs_Height", Convert.ToInt32(Spacing["chart_Abs_Height"] * (CorrelScatter.ChartAreas[0].Position.Height / 100)));
            Spacing.Add("chartArea_Rel_Top", CorrelScatter.ChartAreas[0].Position.Y);
            Spacing.Add("chartArea_Abs_Top", Convert.ToInt32(Spacing["chart_Abs_Height"] * (Spacing["chartArea_Rel_Top"] / 100)));
            Spacing.Add("chartArea_Rel_Bottom", CorrelScatter.ChartAreas[0].Position.Bottom);
            Spacing.Add("chartArea_Abs_Bottom", Convert.ToInt32(Spacing["chart_Abs_Height"] * (Spacing["chartArea_Rel_Bottom"] / 100)));

            Spacing.Add("chartInnerPlot_Abs_Height", Convert.ToInt32(Spacing["chartArea_Abs_Height"] * (CorrelScatter.ChartAreas[0].InnerPlotPosition.Height / 100)));
            Spacing.Add("chartInnerPlot_Rel_Top", CorrelScatter.ChartAreas[0].InnerPlotPosition.Y);
            Spacing.Add("chartInnerPlot_Abs_Top", Convert.ToInt32(Spacing["chartArea_Abs_Height"] * (Spacing["chartInnerPlot_Rel_Top"] / 100)) + Spacing["chartArea_Abs_Top"]);
            Spacing.Add("chartInnerPlot_Rel_Bottom", CorrelScatter.ChartAreas[0].InnerPlotPosition.Bottom);
            Spacing.Add("chartInnerPlot_Abs_Bottom", Convert.ToInt32(Spacing["chartArea_Abs_Height"] * (Spacing["chartInnerPlot_Rel_Bottom"] / 100)) + Spacing["chartArea_Abs_Top"]);

            ConstructLabels();

            HoverLabel_H1.MouseHover += HoverLabel_MouseHoverEvent_H1;
            HoverLabel_H1.MouseLeave += HoverLabel_MouseLeaveEvent_H1;

            HoverLabel_H2.MouseHover += HoverLabel_MouseHoverEvent_H2;
            HoverLabel_H2.MouseLeave += HoverLabel_MouseLeaveEvent_H2;

            HoverLabel_H3.MouseHover += HoverLabel_MouseHoverEvent_H3;
            HoverLabel_H3.MouseLeave += HoverLabel_MouseLeaveEvent_H3;

            HoverLabel_H4.MouseHover += HoverLabel_MouseHoverEvent_H4;
            HoverLabel_H4.MouseLeave += HoverLabel_MouseLeaveEvent_H4;

            HoverLabel_H5.MouseHover += HoverLabel_MouseHoverEvent_H5;
            HoverLabel_H5.MouseLeave += HoverLabel_MouseLeaveEvent_H5;


            HoverLabel_V1.MouseHover += HoverLabel_MouseHoverEvent_V1;
            HoverLabel_V1.MouseLeave += HoverLabel_MouseLeaveEvent_V1;

            HoverLabel_V2.MouseHover += HoverLabel_MouseHoverEvent_V2;
            HoverLabel_V2.MouseLeave += HoverLabel_MouseLeaveEvent_V2;

            HoverLabel_V3.MouseHover += HoverLabel_MouseHoverEvent_V3;
            HoverLabel_V3.MouseLeave += HoverLabel_MouseLeaveEvent_V3;

            HoverLabel_V4.MouseHover += HoverLabel_MouseHoverEvent_V4;
            HoverLabel_V4.MouseLeave += HoverLabel_MouseLeaveEvent_V4;

            HoverLabel_V5.MouseHover += HoverLabel_MouseHoverEvent_V5;
            HoverLabel_V5.MouseLeave += HoverLabel_MouseLeaveEvent_V5;
        }

        private enum QuintantOrientation
        {
            Vertical,
            Horizontal
        }

        private Tuple<double?, double?> GetSubStats(int quintant, QuintantOrientation orientation)
        {
            IEnumerable<DataPoint> pertinentPoints;
            double minBound;
            double maxBound;
            double width;
            double? mean;
            double? stdev;
            if (orientation == QuintantOrientation.Horizontal)
            {
                width = (CorrelDist1.GetMaximum() - CorrelDist1.GetMinimum()) / 5;
                minBound = (quintant - 1) * width + CorrelDist1.GetMinimum();
                maxBound = (quintant) * width + CorrelDist1.GetMinimum();
                pertinentPoints = from DataPoint dp in this.CorrelScatterPoints
                                  where dp.XValue >= minBound && dp.XValue < maxBound
                                  select dp;
                IEnumerable<double> pertinentY = from DataPoint dp in pertinentPoints select dp.YValues.First();
                if (pertinentY.Any())
                {
                    mean = pertinentY.Average();
                    stdev = ExtensionMethods.CalculateStandardDeviation(pertinentY);
                    return new Tuple<double?, double?>(Math.Round((double)mean, 2), Math.Round((double)stdev, 2));
                }
                else
                {
                    return new Tuple<double?, double?>(null, null);
                }
            }
            else if(orientation == QuintantOrientation.Vertical)
            {
                width = (CorrelDist2.GetMaximum() - CorrelDist2.GetMinimum()) / 5;
                minBound = (quintant - 1) * width + CorrelDist2.GetMinimum();
                maxBound = (quintant) * width + CorrelDist2.GetMinimum();
                pertinentPoints = from DataPoint dp in this.CorrelScatterPoints
                                  where dp.YValues.First() >= minBound && dp.YValues.First() < maxBound
                                  select dp;
                IEnumerable<double> pertinentX = from DataPoint dp in pertinentPoints select dp.XValue;
                if (pertinentX.Any())
                {
                    mean = pertinentX.Average();
                    stdev = ExtensionMethods.CalculateStandardDeviation(pertinentX);
                    return new Tuple<double?, double?>(Math.Round((double)mean, 2), Math.Round((double)stdev, 2));
                }
                else
                {
                    mean = null;
                    stdev = null;
                    return new Tuple<double?, double?>(null, null);
                }
            }
            else
            {
                throw new Exception("Unexpected orientation value");
            }
        }


        #region HoverLabel Events

        private void HoverLabel_MouseHoverEvent_H1(object sender, EventArgs e)
        {
            HoverLabel_H1.BackColor = Color.FromArgb(175, 125, 125, 125);
            Tuple<double?, double?> stats = GetSubStats(1, QuintantOrientation.Horizontal);
            HoverLabel_H1.TextAlign = ContentAlignment.MiddleCenter;
            HoverLabel_H1.Text = $"μ: {stats.Item1}\r\nσ: {stats.Item2}";
        }
        private void HoverLabel_MouseLeaveEvent_H1(object sender, EventArgs e)
        {
            HoverLabel_H1.BackColor = Color.FromArgb(25, 125, 125, 125);
            HoverLabel_H1.Text = "";
        }

        private void HoverLabel_MouseHoverEvent_H2(object sender, EventArgs e)
        {
            HoverLabel_H2.BackColor = Color.FromArgb(175, 125, 125, 125);
            Tuple<double?, double?> stats = GetSubStats(2, QuintantOrientation.Horizontal);
            HoverLabel_H2.TextAlign = ContentAlignment.MiddleCenter;
            HoverLabel_H2.Text = $"μ: {stats.Item1}\r\nσ: {stats.Item2}";
        }
        private void HoverLabel_MouseLeaveEvent_H2(object sender, EventArgs e)
        {
            HoverLabel_H2.BackColor = Color.FromArgb(25, 125, 125, 125);
            HoverLabel_H2.Text = "";
        }

        private void HoverLabel_MouseHoverEvent_H3(object sender, EventArgs e)
        {
            HoverLabel_H3.BackColor = Color.FromArgb(175, 125, 125, 125);
            Tuple<double?, double?> stats = GetSubStats(3, QuintantOrientation.Horizontal);
            HoverLabel_H3.TextAlign = ContentAlignment.MiddleCenter;
            HoverLabel_H3.Text = $"μ: {stats.Item1}\r\nσ: {stats.Item2}";
        }
        private void HoverLabel_MouseLeaveEvent_H3(object sender, EventArgs e)
        {
            HoverLabel_H3.BackColor = Color.FromArgb(25, 125, 125, 125);
            HoverLabel_H3.Text = "";
        }

        private void HoverLabel_MouseHoverEvent_H4(object sender, EventArgs e)
        {
            HoverLabel_H4.BackColor = Color.FromArgb(175, 125, 125, 125);
            Tuple<double?, double?> stats = GetSubStats(4, QuintantOrientation.Horizontal);
            HoverLabel_H4.TextAlign = ContentAlignment.MiddleCenter;
            HoverLabel_H4.Text = $"μ: {stats.Item1}\r\nσ: {stats.Item2}";
        }
        private void HoverLabel_MouseLeaveEvent_H4(object sender, EventArgs e)
        {
            HoverLabel_H4.BackColor = Color.FromArgb(25, 125, 125, 125);
            HoverLabel_H4.Text = "";
        }

        private void HoverLabel_MouseHoverEvent_H5(object sender, EventArgs e)
        {
            HoverLabel_H5.BackColor = Color.FromArgb(175, 125, 125, 125);
            Tuple<double?, double?> stats = GetSubStats(5, QuintantOrientation.Horizontal);
            HoverLabel_H5.TextAlign = ContentAlignment.MiddleCenter;
            HoverLabel_H5.Text = $"μ: {stats.Item1}\r\nσ: {stats.Item2}";
            HoverLabel_H5.BringToFront();
        }
        private void HoverLabel_MouseLeaveEvent_H5(object sender, EventArgs e)
        {
            HoverLabel_H5.BackColor = Color.FromArgb(25, 125, 125, 125);
            HoverLabel_H5.Text = "";
        }


        private void HoverLabel_MouseHoverEvent_V1(object sender, EventArgs e)
        {
            HoverLabel_V1.BackColor = Color.FromArgb(175, 125, 125, 125);
            Tuple<double?, double?> stats = GetSubStats(5, QuintantOrientation.Vertical);
            HoverLabel_V1.TextAlign = ContentAlignment.MiddleCenter;
            HoverLabel_V1.Text = $"μ: {stats.Item1}\r\nσ: {stats.Item2}";
        }
        private void HoverLabel_MouseLeaveEvent_V1(object sender, EventArgs e)
        {
            HoverLabel_V1.BackColor = Color.FromArgb(25, 125, 125, 125);
            HoverLabel_V1.Text = "";
        }

        private void HoverLabel_MouseHoverEvent_V2(object sender, EventArgs e)
        {
            HoverLabel_V2.BackColor = Color.FromArgb(175, 125, 125, 125);
            Tuple<double?, double?> stats = GetSubStats(4, QuintantOrientation.Vertical);
            HoverLabel_V2.TextAlign = ContentAlignment.MiddleCenter;
            HoverLabel_V2.Text = $"μ: {stats.Item1}\r\nσ: {stats.Item2}";
        }
        private void HoverLabel_MouseLeaveEvent_V2(object sender, EventArgs e)
        {
            HoverLabel_V2.BackColor = Color.FromArgb(25, 125, 125, 125);
            HoverLabel_V2.Text = "";
        }

        private void HoverLabel_MouseHoverEvent_V3(object sender, EventArgs e)
        {
            HoverLabel_V3.BackColor = Color.FromArgb(175, 125, 125, 125);
            Tuple<double?, double?> stats = GetSubStats(3, QuintantOrientation.Vertical);
            HoverLabel_V3.TextAlign = ContentAlignment.MiddleCenter;
            HoverLabel_V3.Text = $"μ: {stats.Item1}\r\nσ: {stats.Item2}";
        }
        private void HoverLabel_MouseLeaveEvent_V3(object sender, EventArgs e)
        {
            HoverLabel_V3.BackColor = Color.FromArgb(25, 125, 125, 125);
            HoverLabel_V3.Text = "";
        }

        private void HoverLabel_MouseHoverEvent_V4(object sender, EventArgs e)
        {
            HoverLabel_V4.BackColor = Color.FromArgb(175, 125, 125, 125);
            Tuple<double?, double?> stats = GetSubStats(2, QuintantOrientation.Vertical);
            HoverLabel_V4.TextAlign = ContentAlignment.MiddleCenter;
            HoverLabel_V4.Text = $"μ: {stats.Item1}\r\nσ: {stats.Item2}";
        }
        private void HoverLabel_MouseLeaveEvent_V4(object sender, EventArgs e)
        {
            HoverLabel_V4.BackColor = Color.FromArgb(25, 125, 125, 125);
            HoverLabel_V4.Text = "";
        }

        private void HoverLabel_MouseHoverEvent_V5(object sender, EventArgs e)
        {
            HoverLabel_V5.BackColor = Color.FromArgb(175, 125, 125, 125);
            Tuple<double?, double?> stats = GetSubStats(1, QuintantOrientation.Vertical);
            HoverLabel_V5.TextAlign = ContentAlignment.MiddleCenter;
            HoverLabel_V5.Text = $"μ: {stats.Item1}\r\nσ: {stats.Item2}";
            HoverLabel_V5.BringToFront();
        }
        private void HoverLabel_MouseLeaveEvent_V5(object sender, EventArgs e)
        {
            HoverLabel_V5.BackColor = Color.FromArgb(25, 125, 125, 125);
            HoverLabel_V5.Text = "";
        }

        #endregion

        private void XAxisChart_Paint(object sender, PaintEventArgs e)
        {
            //Painting the CorrelationForm calls its subs like this to paint?
            if (!Spacing.ContainsKey("X_MeanPoint_Abs"))
                Spacing.Add("X_MeanPoint_Abs", xAxisChart.ChartAreas[0].AxisX.ValueToPixelPosition(PercentilePoints["X_MeanPoint"].XValue));
            if (!Spacing.ContainsKey("X_LowPoint_Abs"))
                Spacing.Add("X_LowPoint_Abs", xAxisChart.ChartAreas[0].AxisX.ValueToPixelPosition(PercentilePoints["X_LowPoint"].XValue));
            if (!Spacing.ContainsKey("X_HighPoint_Abs"))
                Spacing.Add("X_HighPoint_Abs", xAxisChart.ChartAreas[0].AxisX.ValueToPixelPosition(PercentilePoints["X_HighPoint"].XValue));

        }
        private void YAxisChart_Paint(object sender, PaintEventArgs e)
        {
            //DataPoint lowPoint = PercentilePoints["Y_LowPoint"];
            //double y_low = yAxisChart.ChartAreas[0].AxisX.ValueToPixelPosition(lowPoint.XValue);
            if (!Spacing.ContainsKey("Y_MeanPoint_Abs"))
                Spacing.Add("Y_MeanPoint_Abs", yAxisChart.ChartAreas[0].AxisX.ValueToPixelPosition(PercentilePoints["Y_MeanPoint"].XValue));
            if (!Spacing.ContainsKey("Y_LowPoint_Abs"))
                Spacing.Add("Y_LowPoint_Abs", yAxisChart.ChartAreas[0].AxisX.ValueToPixelPosition(PercentilePoints["Y_LowPoint"].XValue));
            if (!Spacing.ContainsKey("Y_HighPoint_Abs"))
                Spacing.Add("Y_HighPoint_Abs", yAxisChart.ChartAreas[0].AxisX.ValueToPixelPosition(PercentilePoints["Y_HighPoint"].XValue));

            

            if (!RefreshBreak)
            {
                RefreshBreak = true;
                this.Refresh();
            }
        }

        private void CorrelationForm_Paint(object sender, PaintEventArgs e)
        {
            
        }

        private void CorrelScatter_MouseLeave(object sender, EventArgs e)
        {
            if (DrawingMode)
            {
                DrawTool.ResetCursor();
            }
        }

        private void CorrelScatter_MouseEnter(object sender, EventArgs e)
        {
            if (DrawingMode)
            {
                //DrawingTool object set up when you click the button
                DrawTool.EnableCursor();

            }
        }
    }
}
