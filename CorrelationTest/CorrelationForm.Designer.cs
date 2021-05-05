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
            System.Windows.Forms.DataVisualization.Charting.ChartArea chartArea2 = new System.Windows.Forms.DataVisualization.Charting.ChartArea();
            System.Windows.Forms.DataVisualization.Charting.Legend legend2 = new System.Windows.Forms.DataVisualization.Charting.Legend();
            System.Windows.Forms.DataVisualization.Charting.Series series2 = new System.Windows.Forms.DataVisualization.Charting.Series();
            this.CorrelScatter = new System.Windows.Forms.DataVisualization.Charting.Chart();
            this.btn_saveClose = new System.Windows.Forms.Button();
            this.btn_Cancel = new System.Windows.Forms.Button();
            this.numericUpDown_CorrelValue = new System.Windows.Forms.NumericUpDown();
            this.groupBox_CorrelCoef = new System.Windows.Forms.GroupBox();
            this.xAxisChart = new System.Windows.Forms.DataVisualization.Charting.Chart();
            ((System.ComponentModel.ISupportInitialize)(this.CorrelScatter)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDown_CorrelValue)).BeginInit();
            this.groupBox_CorrelCoef.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.xAxisChart)).BeginInit();
            this.SuspendLayout();
            // 
            // CorrelScatter
            // 
            chartArea1.Name = "ChartArea1";
            this.CorrelScatter.ChartAreas.Add(chartArea1);
            legend1.Name = "Legend1";
            this.CorrelScatter.Legends.Add(legend1);
            this.CorrelScatter.Location = new System.Drawing.Point(148, 148);
            this.CorrelScatter.Margin = new System.Windows.Forms.Padding(2);
            this.CorrelScatter.Name = "CorrelScatter";
            series1.ChartArea = "ChartArea1";
            series1.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Point;
            series1.Legend = "Legend1";
            series1.Name = "CorrelSeries";
            series1.XValueType = System.Windows.Forms.DataVisualization.Charting.ChartValueType.Double;
            series1.YValueType = System.Windows.Forms.DataVisualization.Charting.ChartValueType.Double;
            this.CorrelScatter.Series.Add(series1);
            this.CorrelScatter.Size = new System.Drawing.Size(806, 522);
            this.CorrelScatter.TabIndex = 0;
            this.CorrelScatter.Text = "Scatterplot";
            // 
            // btn_saveClose
            // 
            this.btn_saveClose.Location = new System.Drawing.Point(804, 552);
            this.btn_saveClose.Margin = new System.Windows.Forms.Padding(2);
            this.btn_saveClose.Name = "btn_saveClose";
            this.btn_saveClose.Size = new System.Drawing.Size(138, 31);
            this.btn_saveClose.TabIndex = 1;
            this.btn_saveClose.Text = "Save";
            this.btn_saveClose.UseVisualStyleBackColor = true;
            this.btn_saveClose.Click += new System.EventHandler(this.btn_saveClose_Click);
            // 
            // btn_Cancel
            // 
            this.btn_Cancel.Location = new System.Drawing.Point(804, 587);
            this.btn_Cancel.Margin = new System.Windows.Forms.Padding(2);
            this.btn_Cancel.Name = "btn_Cancel";
            this.btn_Cancel.Size = new System.Drawing.Size(138, 31);
            this.btn_Cancel.TabIndex = 2;
            this.btn_Cancel.Text = "Cancel";
            this.btn_Cancel.UseVisualStyleBackColor = true;
            this.btn_Cancel.Click += new System.EventHandler(this.btn_Cancel_Click);
            // 
            // numericUpDown_CorrelValue
            // 
            this.numericUpDown_CorrelValue.Font = new System.Drawing.Font("Microsoft Sans Serif", 13.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.numericUpDown_CorrelValue.Location = new System.Drawing.Point(15, 17);
            this.numericUpDown_CorrelValue.Margin = new System.Windows.Forms.Padding(2);
            this.numericUpDown_CorrelValue.Name = "numericUpDown_CorrelValue";
            this.numericUpDown_CorrelValue.Size = new System.Drawing.Size(107, 28);
            this.numericUpDown_CorrelValue.TabIndex = 3;
            this.numericUpDown_CorrelValue.ValueChanged += new System.EventHandler(this.numericUpDown_CorrelValue_ValueChanged);
            this.numericUpDown_CorrelValue.MouseDown += new System.Windows.Forms.MouseEventHandler(this.numericUpDown_CorrelValue_MouseDown);
            // 
            // groupBox_CorrelCoef
            // 
            this.groupBox_CorrelCoef.BackColor = System.Drawing.SystemColors.Info;
            this.groupBox_CorrelCoef.Controls.Add(this.numericUpDown_CorrelValue);
            this.groupBox_CorrelCoef.Location = new System.Drawing.Point(804, 461);
            this.groupBox_CorrelCoef.Margin = new System.Windows.Forms.Padding(2);
            this.groupBox_CorrelCoef.Name = "groupBox_CorrelCoef";
            this.groupBox_CorrelCoef.Padding = new System.Windows.Forms.Padding(2);
            this.groupBox_CorrelCoef.Size = new System.Drawing.Size(138, 86);
            this.groupBox_CorrelCoef.TabIndex = 5;
            this.groupBox_CorrelCoef.TabStop = false;
            this.groupBox_CorrelCoef.Text = "Correlation Coefficient";
            this.groupBox_CorrelCoef.Enter += new System.EventHandler(this.groupBox_CorrelCoef_Enter);
            // 
            // xAxisChart
            // 
            chartArea2.Name = "ChartArea1";
            this.xAxisChart.ChartAreas.Add(chartArea2);
            legend2.Name = "Legend1";
            this.xAxisChart.Legends.Add(legend2);
            this.xAxisChart.Location = new System.Drawing.Point(222, 56);
            this.xAxisChart.Name = "xAxisChart";
            series2.ChartArea = "ChartArea1";
            series2.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Area;
            series2.Legend = "Legend1";
            series2.Name = "Series1";
            this.xAxisChart.Series.Add(series2);
            this.xAxisChart.Size = new System.Drawing.Size(564, 105);
            this.xAxisChart.TabIndex = 6;
            this.xAxisChart.Text = "chart1";
            // 
            // CorrelationForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.Window;
            this.ClientSize = new System.Drawing.Size(968, 681);
            this.Controls.Add(this.xAxisChart);
            this.Controls.Add(this.groupBox_CorrelCoef);
            this.Controls.Add(this.btn_Cancel);
            this.Controls.Add(this.btn_saveClose);
            this.Controls.Add(this.CorrelScatter);
            this.Margin = new System.Windows.Forms.Padding(2);
            this.Name = "CorrelationForm";
            this.Text = "CorrelationForm";
            this.Load += new System.EventHandler(this.CorrelationForm_Load);
            ((System.ComponentModel.ISupportInitialize)(this.CorrelScatter)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDown_CorrelValue)).EndInit();
            this.groupBox_CorrelCoef.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.xAxisChart)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.DataVisualization.Charting.Chart CorrelScatter;
        private System.Windows.Forms.Button btn_saveClose;
        private System.Windows.Forms.Button btn_Cancel;
        private System.Windows.Forms.NumericUpDown numericUpDown_CorrelValue;
        private System.Windows.Forms.GroupBox groupBox_CorrelCoef;
        private System.Windows.Forms.DataVisualization.Charting.Chart xAxisChart;
    }
}