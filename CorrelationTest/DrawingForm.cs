using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CorrelationTest
{
    public partial class DrawingForm : Form
    {
        LineSegment NewSegment { get; set; } = new LineSegment();

        public DrawingForm()
        {
            InitializeComponent();
        }

        private void DrawingForm_Paint(object sender, PaintEventArgs e)
        {
            if (NewSegment != null)
            {
                if (NewSegment.Points.Count() > 1)
                {
                    Pen pen = new Pen(Color.FromArgb(255, 0, 0, 0));
                    e.Graphics.DrawLines(pen, NewSegment.GetPoints());
                }
            }
        }

        private void DrawingForm_MouseClick(object sender, MouseEventArgs e)
        {
            if (NewSegment.AddPoint(e.Location))
            {
                if (NewSegment.Points.Count() > 1)
                {
                    Refresh();
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            MessageBox.Show(this.NewSegment.GetSlope().ToString());
            MessageBox.Show(this.NewSegment.GetCorrelation().ToString());
        }
    }
}
