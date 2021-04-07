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
            //Create & set example points
            for (int i = 0; i < 20; i++)
            {
                double input = ((double)i + 
                    1) / 100;
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
        }

        private void CorrelScatter_Click(object sender, EventArgs e)
        {

        }
    }
}
