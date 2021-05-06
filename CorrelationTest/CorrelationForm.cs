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
            Transitivity,
            PSD,
            Conformal
        }
        private List<DataPoint> CorrelScatterPoints { get; set; } = new List<DataPoint>();
        private Color existingColor { get; set; }
        private Color existingColor_Markers { get; set; }
        private int helperStage { get; set; }
        private TextBox textboxMinimum = new TextBox();
        private TextBox textboxMidpoint = new TextBox();
        private TextBox textboxMaximum = new TextBox();
        private DrawnCorrelation DrawnCorrel { get; set; }
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
            this.CorrelScatter = new System.Windows.Forms.DataVisualization.Charting.Chart();
            this.CorrelScatter.Series["CorrelSeries"].MarkerStyle = MarkerStyle.Circle;
            //Set the axis scale
            //this.CorrelScatter.ChartAreas[0].AxisX.LabelStyle.Format = "0.00";
            //this.CorrelScatter.ChartAreas[0].AxisY.LabelStyle.Format = "0.00";
            this.CorrelScatter.ChartAreas[0].AxisY.Enabled = AxisEnabled.False;
            this.CorrelScatter.ChartAreas[0].AxisY2.Enabled = AxisEnabled.True;
            this.CorrelScatter.ChartAreas[0].AxisX.Interval = .5;
            this.CorrelScatter.ChartAreas[0].AxisX.Minimum = CorrelDist1.GetMinimum_X();
            this.CorrelScatter.ChartAreas[0].AxisX.Maximum = CorrelDist1.GetMaximum_X();
            this.CorrelScatter.ChartAreas[0].AxisY2.Interval = .5;
            this.CorrelScatter.ChartAreas[0].AxisY2.Minimum = CorrelDist2.GetMinimum_X();
            this.CorrelScatter.ChartAreas[0].AxisY2.Maximum = CorrelDist2.GetMaximum_X();

            foreach(DataPoint dp in CorrelScatterPoints)
            {
                this.CorrelScatter.Series["CorrelSeries"].Points.AddXY(dp.XValue, dp.YValues[0]);
            }
        }

        private void CorrelationForm_Load(object sender, EventArgs e)
        {
            Sheets.CorrelationSheet CorrelSheet = Sheets.CorrelationSheet.ConstructFromXlCorrelationSheet();
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
            //this.CorrelScatter.ChartAreas[0].AxisX.LabelStyle.Format = "0.00";
            //this.CorrelScatter.ChartAreas[0].AxisY.LabelStyle.Format = "0.00";
            this.CorrelScatter.ChartAreas[0].AxisY.Enabled = AxisEnabled.False;
            this.CorrelScatter.ChartAreas[0].AxisY2.Enabled = AxisEnabled.True;
            this.CorrelScatter.ChartAreas[0].AxisX.Interval = .5;
            this.CorrelScatter.ChartAreas[0].AxisX.Minimum = CorrelDist1.GetMinimum_X();
            this.CorrelScatter.ChartAreas[0].AxisX.Maximum = CorrelDist1.GetMaximum_X();
            this.CorrelScatter.ChartAreas[0].AxisY2.Interval = .5;
            this.CorrelScatter.ChartAreas[0].AxisY2.Minimum = CorrelDist2.GetMinimum_X();
            this.CorrelScatter.ChartAreas[0].AxisY2.Maximum = CorrelDist2.GetMaximum_X();

            

            //double xMin = (from DataPoint dp in this.CorrelScatter.Series["CorrelSeries"].Points select dp.XValue).Min();
            //double xMax = (from DataPoint dp in this.CorrelScatter.Series["CorrelSeries"].Points select dp.XValue).Max();
            //this.CorrelScatter.ChartAreas[0].AxisX.Minimum = Math.Floor(xMin);
            //this.CorrelScatter.ChartAreas[0].AxisX.Maximum = Math.Ceiling(xMax);

            //double yMin = (from DataPoint dp in this.CorrelScatter.Series["CorrelSeries"].Points select dp.YValues.First()).Min();
            //double yMax = (from DataPoint dp in this.CorrelScatter.Series["CorrelSeries"].Points select dp.YValues.First()).Max();
            //this.CorrelScatter.ChartAreas[0].AxisY.Minimum = Math.Floor(yMin);
            //this.CorrelScatter.ChartAreas[0].AxisY.Maximum = Math.Ceiling(yMax);

            Excel.Range xlSelection = ThisAddIn.MyApp.Selection;
            int index1 = xlSelection.Row - (CorrelSheet.xlMatrixCell.Row + 1);
            int index2 = xlSelection.Column - CorrelSheet.xlMatrixCell.Column;


            Tuple<double, double> trans_bounds = CorrelSheet.CorrelMatrix.GetTransitivityBounds(index1, index2);  //<min, max>
            numericUpDown_CorrelValue.TextAlign = HorizontalAlignment.Center;
            numericUpDown_CorrelValue.DecimalPlaces = 2;
            numericUpDown_CorrelValue.Minimum = Convert.ToDecimal(trans_bounds.Item1);
            numericUpDown_CorrelValue.Maximum = Convert.ToDecimal(trans_bounds.Item2);
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
            else if (existingValue < trans_bounds.Item1)
            {
                numericUpDown_CorrelValue.Value = Convert.ToDecimal(trans_bounds.Item1);  //Set to min
                if(trans_bounds.Item1 == -1)
                    CoefficientBox_FlagError(CoefficientBox_ErrorType.Conformal);
                else
                    CoefficientBox_FlagError(CoefficientBox_ErrorType.Transitivity);
                lastValue = numericUpDown_CorrelValue.Value;
            }
            else if (existingValue > trans_bounds.Item2)
            {
                numericUpDown_CorrelValue.Value = Convert.ToDecimal(trans_bounds.Item2);  //Set to max
                if (trans_bounds.Item1 == 1)
                    CoefficientBox_FlagError(CoefficientBox_ErrorType.Conformal);
                else
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
        }

        private void LoadYAxisDistribution()
        {
            //Build a series off the distribution
            System.IO.MemoryStream myStream = new System.IO.MemoryStream();
            Chart yAxisChart = new Chart();
            xAxisChart.Serializer.Save(myStream);
            yAxisChart.Serializer.Load(myStream);

            yAxisChart.Series.Clear();
            Series Series1 = new Series();
            yAxisChart.Series.Add(Series1);
            yAxisChart.Series["Series1"].ChartType = SeriesChartType.Bar;
            yAxisChart.Width = xAxisChart.Height;
            yAxisChart.Left = CorrelScatter.Left - yAxisChart.Width;
            yAxisChart.Top = CorrelScatter.Top + 12;
            yAxisChart.Height = CorrelScatter.Height - 10;
            yAxisChart.Series["Series1"].YValuesPerPoint = 1;
            yAxisChart.ChartAreas[0].AxisX.Interval = 0.5;
            yAxisChart.Series["Series1"].IsVisibleInLegend = false;
            yAxisChart.Series["Series1"]["PixelPointWidth"] = "2";

            int steps = 1000;
            double minimum = CorrelDist2.GetMinimum_X();
            double maximum = CorrelDist2.GetMaximum_X();
            double step = (maximum - minimum) / steps;

            yAxisChart.ChartAreas[0].AxisX.Minimum = minimum;
            yAxisChart.ChartAreas[0].AxisX.Maximum = maximum;

            for (int i = 0; i < steps; i++)
            {                
                double x = minimum + step * i;
                double y = CorrelDist2.GetPDF_Value(x);
                yAxisChart.Series["Series1"].Points.AddXY(x, y);
            }
            //yAxisChart.ChartAreas["ChartArea1"].Area3DStyle.Rotation = 45;
            //yAxisChart.ChartAreas["ChartArea1"].Area3DStyle.Inclination = 45;
            yAxisChart.ChartAreas[0].AxisY.IntervalAutoMode = IntervalAutoMode.VariableCount;
            yAxisChart.ChartAreas[0].AxisX.IntervalAutoMode = IntervalAutoMode.VariableCount;
            yAxisChart.ChartAreas[0].AxisY.Interval = 0.1;
            yAxisChart.ChartAreas[0].AxisY.IsReversed = true;
            
            this.Controls.Add(yAxisChart);
        }

        private void LoadXAxisDistribution()
        {
            //Build a series off the distribution
            this.xAxisChart.Left = CorrelScatter.Left - 40;
            this.xAxisChart.Top = CorrelScatter.Top - 150;
            this.xAxisChart.Height = 150;
            this.xAxisChart.Width = CorrelScatter.Width - 15;
            this.xAxisChart.Series["Series1"].YValuesPerPoint = 1;
            this.xAxisChart.ChartAreas[0].AxisX.Interval = 0.5;

            int steps = 100;
            double minimum = CorrelDist1.GetMinimum_X();
            double maximum = CorrelDist1.GetMaximum_X();
            double step = (maximum - minimum) / steps;

            this.xAxisChart.ChartAreas[0].AxisX.Minimum = minimum;
            this.xAxisChart.ChartAreas[0].AxisX.Maximum = maximum;

            for (int i = 0; i < steps; i++)
            {
                double x = minimum + step * i;
                double y = CorrelDist1.GetPDF_Value(x);
                xAxisChart.Series["Series1"].Points.AddXY(x, y);
            }
            xAxisChart.ChartAreas[0].AxisY.IntervalAutoMode = IntervalAutoMode.VariableCount;
            xAxisChart.ChartAreas[0].AxisX.IntervalAutoMode = IntervalAutoMode.VariableCount;
            xAxisChart.ChartAreas[0].AxisX.LabelStyle.Enabled = false;
        }

        public void CoefficientBox_Reset()
        {
            //DEFAULT "Info" yellow
            groupBox_CorrelCoef.BackColor = Color.FromArgb(255, 255, 225);
            this.label_coefErrors.Text = "No Errors";
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
            if (lastValue == numericUpDown_CorrelValue.Value)
            {
                if (numericUpDown_CorrelValue.Value >= numericUpDown_CorrelValue.Maximum)
                {
                    if (numericUpDown_CorrelValue.Maximum == 1)
                        CoefficientBox_FlagError(CoefficientBox_ErrorType.Conformal);
                    else
                        CoefficientBox_FlagError(CoefficientBox_ErrorType.Transitivity);
                }
                else if (numericUpDown_CorrelValue.Value <= numericUpDown_CorrelValue.Minimum)
                {
                    if (numericUpDown_CorrelValue.Minimum == -1)
                        CoefficientBox_FlagError(CoefficientBox_ErrorType.Conformal);
                    else
                        CoefficientBox_FlagError(CoefficientBox_ErrorType.Transitivity);
                }
            }
            lastValue = numericUpDown_CorrelValue.Value;
        }

        private void numericUpDown_CorrelValue_ValueChanged(object sender, EventArgs e)
        {
            //Save the old value
            CoefficientBox_Reset();
        }

        private void SetupHelper()
        {
            int vertical = CorrelScatter.Height / 2;
            int min = CorrelScatter.Width / 4;
            int mid = min * 2;
            int max = min * 3;

            textboxMinimum.Top = vertical;
            textboxMinimum.Left = min;
            textboxMinimum.Height = 50;
            textboxMinimum.Width = 50;
            textboxMidpoint.Top = vertical;
            textboxMidpoint.Left = mid;
            textboxMidpoint.Height = 50;
            textboxMidpoint.Width = 50;
            textboxMaximum.Top = vertical;
            textboxMaximum.Left = max;
            textboxMaximum.Height = 50;
            textboxMaximum.Width = 50;

        }

        private void btn_LaunchHelper_Click(object sender, EventArgs e)
        {
            if(helperStage == 0)
            {
                //Load minimum
                existingColor = CorrelScatter.ChartAreas[0].BackColor;
                CorrelScatter.ChartAreas[0].BackColor = Color.FromArgb(195, 195, 195);
                this.btn_LaunchHelper.Text = ">> Next >>";
                this.Controls.Add(textboxMinimum);
                textboxMinimum.BringToFront();
                labelHelper.AutoSize = true;
                labelHelper.Top = textboxMinimum.Top - 50;
                labelHelper.Left = textboxMinimum.Left;
                labelHelper.Text = $"If X is {CorrelDist1.GetMinimum_X()}, what do you expect Y to be?";
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
                labelHelper.Text = $"If X is {(CorrelDist1.GetMaximum_X() - CorrelDist1.GetMinimum_X()) / 2}, what do you expect Y to be?";
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
                labelHelper.Text = $"If X is {CorrelDist1.GetMaximum_X()}, what do you expect Y to be?";
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

                helperStage = 0;
            }
        }

        private void btn_LaunchDrawCorrelation_Click(object sender, EventArgs e)
        {
            if (!DrawingMode)
            {
                //Turn on DrawingMode
                DrawingMode = true;
                this.DrawnCorrel = new DrawnCorrelation();
                //Disable the other buttons
                this.btn_LaunchHelper.Enabled = false;
                this.btn_saveClose.Enabled = false;
                this.UpDownEnabled = this.numericUpDown_CorrelValue.Enabled;
                if(this.UpDownEnabled)
                    this.numericUpDown_CorrelValue.Enabled = false;
                btn_LaunchDrawCorrelation.Text = "Done Drawing";
                existingColor = CorrelScatter.ChartAreas[0].BackColor;
                CorrelScatter.ChartAreas[0].BackColor = Color.FromArgb(195, 195, 195);
                existingColor_Markers = CorrelScatter.Series["CorrelSeries"].Color;
                CorrelScatter.Series["CorrelSeries"].Color = Color.FromArgb(0, 195, 195, 195);
            }
            else
            {
                //Turn off DrawingMode
                this.DrawnCorrel = null;
                this.btn_LaunchHelper.Enabled = true;
                this.btn_saveClose.Enabled = true;
                this.numericUpDown_CorrelValue.Enabled = this.UpDownEnabled;    //Reset to original state
                btn_LaunchDrawCorrelation.Text = "Draw Correlation";
                CorrelScatter.ChartAreas[0].BackColor = existingColor;
                CorrelScatter.Series["CorrelSeries"].Color = existingColor_Markers;
                DrawingMode = false;
            }
        }

        private void CorrelScatter_MouseClick(object sender, MouseEventArgs e)
        {
            if (!DrawingMode)
                return;
            if (DrawnCorrel.AddPoint(e.Location))
            {
                if (DrawnCorrel.Points.Count() > 1)
                {
                    Refresh();
                }
            }
        }

        private void CorrelScatter_Paint(object sender, PaintEventArgs e)
        {
            if (DrawingMode)
            {
                if (DrawnCorrel.Points.Count() > 1)
                {
                    Pen pen = new Pen(Color.FromArgb(255, 0, 0, 0));
                    e.Graphics.DrawLines(pen, DrawnCorrel.GetPoints());
                }
                else if(DrawnCorrel == null)
                {
                    this.ReloadCorrelScatter();
                }
            }
        }
    }
}
