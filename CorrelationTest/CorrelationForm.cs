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
        private Distribution CorrelDist1 { get; set; }
        private Distribution CorrelDist2 { get; set; }
        public CorrelationForm(Distribution correlDist1, Distribution correlDist2)
        {
            this.CorrelDist1 = correlDist1;
            this.CorrelDist2 = correlDist2;
            InitializeComponent();
        }

        private void CorrelationForm_Load(object sender, EventArgs e)
        {
            for (int i = 0; i < 20; i++)
            {
                double input = ((double)i + 1) / 100;
                double x = this.CorrelDist1.DistributionObj.InverseDistributionFunction(input);
                double y = this.CorrelDist2.DistributionObj.InverseDistributionFunction(input);
                this.CorrelScatter.Series["CorrelSeries"].Points.AddXY(x, y);
            }
            
        }

        private void CorrelScatter_Click(object sender, EventArgs e)
        {

        }
    }
}
