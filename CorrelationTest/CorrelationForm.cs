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
        private IEstimateDistribution CorrelDist1 { get; set; }
        private IEstimateDistribution CorrelDist2 { get; set; }
        public CorrelationForm(IEstimateDistribution correlDist1, IEstimateDistribution correlDist2)
        {
            this.CorrelDist1 = correlDist1;
            this.CorrelDist2 = correlDist2;
            InitializeComponent();
        }

        private void CorrelationForm_Load(object sender, EventArgs e)
        {
            Sheets.CorrelationSheet CorrelSheet = Sheets.CorrelationSheet.ConstructFromXlCorrelationSheet();

            //Create & set example points
            for (int i = 0; i < 20; i++)
            {
                double input = ((double)i + 1) / 100;
                double x = CorrelDist1.GetInverse(input);
                double y = CorrelDist2.GetInverse(input);
                this.CorrelScatter.Series["CorrelSeries"].Points.AddXY(x, y);
            }
            //Set the axis scale
            this.CorrelScatter.ChartAreas[0].AxisX.LabelStyle.Format = "0.00";
            this.CorrelScatter.ChartAreas[0].AxisY.LabelStyle.Format = "0.00";
            this.CorrelScatter.ChartAreas[0].AxisX.Interval = .5;
            this.CorrelScatter.ChartAreas[0].AxisY.Interval = .5;

            double xMin = (from DataPoint dp in this.CorrelScatter.Series["CorrelSeries"].Points select dp.XValue).Min();
            double xMax = (from DataPoint dp in this.CorrelScatter.Series["CorrelSeries"].Points select dp.XValue).Max();
            this.CorrelScatter.ChartAreas[0].AxisX.Minimum = Math.Floor(xMin);
            this.CorrelScatter.ChartAreas[0].AxisX.Maximum = Math.Ceiling(xMax);

            double yMin = (from DataPoint dp in this.CorrelScatter.Series["CorrelSeries"].Points select dp.YValues.First()).Min();
            double yMax = (from DataPoint dp in this.CorrelScatter.Series["CorrelSeries"].Points select dp.YValues.First()).Max();
            this.CorrelScatter.ChartAreas[0].AxisY.Minimum = Math.Floor(yMin);
            this.CorrelScatter.ChartAreas[0].AxisY.Maximum = Math.Ceiling(yMax);

            Excel.Range xlSelection = ThisAddIn.MyApp.Selection;
            int index1 = xlSelection.Row - (CorrelSheet.xlMatrixCell.Row + 1);
            int index2 = xlSelection.Column - CorrelSheet.xlMatrixCell.Column;


            Tuple<double, double> trans_bounds = CorrelSheet.CorrelMatrix.GetTransitivityBounds(index1, index2);  //<min, max>
            numericUpDown_CorrelValue.TextAlign = HorizontalAlignment.Center;
            numericUpDown_CorrelValue.DecimalPlaces = 2;
            numericUpDown_CorrelValue.Minimum = Convert.ToDecimal(trans_bounds.Item1);
            numericUpDown_CorrelValue.Maximum = Convert.ToDecimal(trans_bounds.Item2);
            numericUpDown_CorrelValue.Increment = Convert.ToDecimal(0.05);

            
            double existingValue = Convert.ToDouble(ThisAddIn.MyApp.Selection.value);
            if (existingValue >= trans_bounds.Item1 && existingValue <= trans_bounds.Item2)
                numericUpDown_CorrelValue.Value = Convert.ToDecimal(existingValue); //Keep existing matrix value
            else if(existingValue < trans_bounds.Item1)
                numericUpDown_CorrelValue.Value = Convert.ToDecimal(trans_bounds.Item1);  //Set to min
            else if(existingValue > trans_bounds.Item2)
                numericUpDown_CorrelValue.Value = Convert.ToDecimal(trans_bounds.Item2);  //Set to max
        }

        private void CorrelScatter_Click(object sender, EventArgs e)
        {

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
    }
}
