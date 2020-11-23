namespace CorrelationTest
{
    partial class CorrelationForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.Windows.Forms.DataVisualization.Charting.ChartArea chartArea1 = new System.Windows.Forms.DataVisualization.Charting.ChartArea();
            System.Windows.Forms.DataVisualization.Charting.Legend legend1 = new System.Windows.Forms.DataVisualization.Charting.Legend();
            System.Windows.Forms.DataVisualization.Charting.Series series1 = new System.Windows.Forms.DataVisualization.Charting.Series();
            this.CorrelScatter = new System.Windows.Forms.DataVisualization.Charting.Chart();
            ((System.ComponentModel.ISupportInitialize)(this.CorrelScatter)).BeginInit();
            this.SuspendLayout();
            // 
            // CorrelScatter
            // 
            chartArea1.Name = "ChartArea1";
            this.CorrelScatter.ChartAreas.Add(chartArea1);
            legend1.Name = "Legend1";
            this.CorrelScatter.Legends.Add(legend1);
            this.CorrelScatter.Location = new System.Drawing.Point(12, 12);
            this.CorrelScatter.Name = "CorrelScatter";
            series1.ChartArea = "ChartArea1";
            series1.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Point;
            series1.Legend = "Legend1";
            series1.Name = "CorrelSeries";
            series1.XValueType = System.Windows.Forms.DataVisualization.Charting.ChartValueType.Double;
            series1.YValueType = System.Windows.Forms.DataVisualization.Charting.ChartValueType.Double;
            this.CorrelScatter.Series.Add(series1);
            this.CorrelScatter.Size = new System.Drawing.Size(1072, 673);
            this.CorrelScatter.TabIndex = 0;
            this.CorrelScatter.Text = "Scatterplot";
            this.CorrelScatter.Click += new System.EventHandler(this.CorrelScatter_Click);
            // 
            // CorrelationForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1205, 745);
            this.Controls.Add(this.CorrelScatter);
            this.Name = "CorrelationForm";
            this.Text = "CorrelationForm";
            this.Load += new System.EventHandler(this.CorrelationForm_Load);
            ((System.ComponentModel.ISupportInitialize)(this.CorrelScatter)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.DataVisualization.Charting.Chart CorrelScatter;
    }
}