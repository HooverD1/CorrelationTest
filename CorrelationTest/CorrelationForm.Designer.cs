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
            System.Windows.Forms.DataVisualization.Charting.ChartArea chartArea2 = new System.Windows.Forms.DataVisualization.Charting.ChartArea();
            System.Windows.Forms.DataVisualization.Charting.Series series1 = new System.Windows.Forms.DataVisualization.Charting.Series();
            this.CorrelScatter = new System.Windows.Forms.DataVisualization.Charting.Chart();
            this.btn_saveClose = new System.Windows.Forms.Button();
            this.btn_Cancel = new System.Windows.Forms.Button();
            this.numericUpDown_CorrelValue = new System.Windows.Forms.NumericUpDown();
            this.groupBox_CorrelCoef = new System.Windows.Forms.GroupBox();
            this.xAxisChart = new System.Windows.Forms.DataVisualization.Charting.Chart();
            this.btn_LaunchHelper = new System.Windows.Forms.Button();
            this.btn_LaunchDrawCorrelation = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.CorrelScatter)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDown_CorrelValue)).BeginInit();
            this.groupBox_CorrelCoef.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.xAxisChart)).BeginInit();
            this.SuspendLayout();
            // 
            // CorrelScatter
            // 
            chartArea1.AxisX.ScaleView.Zoomable = false;
            chartArea1.AxisX2.ScaleView.Zoomable = false;
            chartArea1.AxisY.Enabled = System.Windows.Forms.DataVisualization.Charting.AxisEnabled.False;
            chartArea1.AxisY.ScaleView.Zoomable = false;
            chartArea1.AxisY2.Enabled = System.Windows.Forms.DataVisualization.Charting.AxisEnabled.True;
            chartArea1.AxisY2.ScaleView.Zoomable = false;
            chartArea1.Name = "ChartArea1";
            chartArea1.Position.Auto = false;
            chartArea1.Position.Height = 94F;
            chartArea1.Position.Width = 94F;
            chartArea1.Position.X = 3F;
            chartArea1.Position.Y = 3F;
            this.CorrelScatter.ChartAreas.Add(chartArea1);
            this.CorrelScatter.Location = new System.Drawing.Point(197, 182);
            this.CorrelScatter.Margin = new System.Windows.Forms.Padding(0);
            this.CorrelScatter.Name = "CorrelScatter";
            this.CorrelScatter.Size = new System.Drawing.Size(750, 750);
            this.CorrelScatter.TabIndex = 0;
            this.CorrelScatter.Text = "Scatterplot";
            this.CorrelScatter.CursorPositionChanged += new System.EventHandler<System.Windows.Forms.DataVisualization.Charting.CursorEventArgs>(this.CorrelScatter_CursorPositionChanged);
            this.CorrelScatter.Paint += new System.Windows.Forms.PaintEventHandler(this.CorrelScatter_Paint);
            this.CorrelScatter.KeyDown += new System.Windows.Forms.KeyEventHandler(this.CorrelScatter_KeyDown);
            this.CorrelScatter.MouseClick += new System.Windows.Forms.MouseEventHandler(this.CorrelScatter_MouseClick);
            this.CorrelScatter.MouseDown += new System.Windows.Forms.MouseEventHandler(this.CorrelScatter_MouseDown);
            this.CorrelScatter.MouseEnter += new System.EventHandler(this.CorrelScatter_MouseEnter);
            this.CorrelScatter.MouseLeave += new System.EventHandler(this.CorrelScatter_MouseLeave);
            this.CorrelScatter.MouseUp += new System.Windows.Forms.MouseEventHandler(this.CorrelScatter_MouseUp);
            // 
            // btn_saveClose
            // 
            this.btn_saveClose.Location = new System.Drawing.Point(968, 790);
            this.btn_saveClose.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.btn_saveClose.Name = "btn_saveClose";
            this.btn_saveClose.Size = new System.Drawing.Size(184, 38);
            this.btn_saveClose.TabIndex = 1;
            this.btn_saveClose.Text = "Save";
            this.btn_saveClose.UseVisualStyleBackColor = true;
            this.btn_saveClose.Click += new System.EventHandler(this.btn_saveClose_Click);
            // 
            // btn_Cancel
            // 
            this.btn_Cancel.Location = new System.Drawing.Point(968, 833);
            this.btn_Cancel.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.btn_Cancel.Name = "btn_Cancel";
            this.btn_Cancel.Size = new System.Drawing.Size(184, 38);
            this.btn_Cancel.TabIndex = 2;
            this.btn_Cancel.Text = "Cancel";
            this.btn_Cancel.UseVisualStyleBackColor = true;
            this.btn_Cancel.Click += new System.EventHandler(this.btn_Cancel_Click);
            // 
            // numericUpDown_CorrelValue
            // 
            this.numericUpDown_CorrelValue.Font = new System.Drawing.Font("Microsoft Sans Serif", 13.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.numericUpDown_CorrelValue.Location = new System.Drawing.Point(20, 21);
            this.numericUpDown_CorrelValue.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.numericUpDown_CorrelValue.Name = "numericUpDown_CorrelValue";
            this.numericUpDown_CorrelValue.Size = new System.Drawing.Size(143, 34);
            this.numericUpDown_CorrelValue.TabIndex = 3;
            this.numericUpDown_CorrelValue.ValueChanged += new System.EventHandler(this.numericUpDown_CorrelValue_ValueChanged);
            this.numericUpDown_CorrelValue.MouseDown += new System.Windows.Forms.MouseEventHandler(this.numericUpDown_CorrelValue_MouseDown);
            this.numericUpDown_CorrelValue.MouseUp += new System.Windows.Forms.MouseEventHandler(this.numericUpDown_CorrelValue_MouseUp);
            // 
            // groupBox_CorrelCoef
            // 
            this.groupBox_CorrelCoef.BackColor = System.Drawing.SystemColors.Info;
            this.groupBox_CorrelCoef.Controls.Add(this.numericUpDown_CorrelValue);
            this.groupBox_CorrelCoef.Location = new System.Drawing.Point(968, 678);
            this.groupBox_CorrelCoef.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.groupBox_CorrelCoef.Name = "groupBox_CorrelCoef";
            this.groupBox_CorrelCoef.Padding = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.groupBox_CorrelCoef.Size = new System.Drawing.Size(184, 106);
            this.groupBox_CorrelCoef.TabIndex = 5;
            this.groupBox_CorrelCoef.TabStop = false;
            this.groupBox_CorrelCoef.Text = "Correlation Coefficient";
            this.groupBox_CorrelCoef.Enter += new System.EventHandler(this.groupBox_CorrelCoef_Enter);
            // 
            // xAxisChart
            // 
            chartArea2.Name = "ChartArea1";
            this.xAxisChart.ChartAreas.Add(chartArea2);
            this.xAxisChart.Location = new System.Drawing.Point(226, 63);
            this.xAxisChart.Margin = new System.Windows.Forms.Padding(0);
            this.xAxisChart.Name = "xAxisChart";
            series1.ChartArea = "ChartArea1";
            series1.Name = "Series1";
            this.xAxisChart.Series.Add(series1);
            this.xAxisChart.Size = new System.Drawing.Size(752, 129);
            this.xAxisChart.TabIndex = 6;
            this.xAxisChart.Text = "chart1";
            // 
            // btn_LaunchHelper
            // 
            this.btn_LaunchHelper.Location = new System.Drawing.Point(968, 613);
            this.btn_LaunchHelper.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.btn_LaunchHelper.Name = "btn_LaunchHelper";
            this.btn_LaunchHelper.Size = new System.Drawing.Size(184, 38);
            this.btn_LaunchHelper.TabIndex = 7;
            this.btn_LaunchHelper.Text = "Use Guided Correlation";
            this.btn_LaunchHelper.UseVisualStyleBackColor = true;
            this.btn_LaunchHelper.Click += new System.EventHandler(this.btn_LaunchHelper_Click);
            // 
            // btn_LaunchDrawCorrelation
            // 
            this.btn_LaunchDrawCorrelation.Location = new System.Drawing.Point(968, 571);
            this.btn_LaunchDrawCorrelation.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.btn_LaunchDrawCorrelation.Name = "btn_LaunchDrawCorrelation";
            this.btn_LaunchDrawCorrelation.Size = new System.Drawing.Size(184, 38);
            this.btn_LaunchDrawCorrelation.TabIndex = 8;
            this.btn_LaunchDrawCorrelation.Text = "Draw Correlation";
            this.btn_LaunchDrawCorrelation.UseVisualStyleBackColor = true;
            this.btn_LaunchDrawCorrelation.Click += new System.EventHandler(this.btn_LaunchDrawCorrelation_Click);
            // 
            // CorrelationForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.Window;
            this.ClientSize = new System.Drawing.Size(1175, 932);
            this.Controls.Add(this.btn_LaunchDrawCorrelation);
            this.Controls.Add(this.btn_LaunchHelper);
            this.Controls.Add(this.xAxisChart);
            this.Controls.Add(this.groupBox_CorrelCoef);
            this.Controls.Add(this.btn_Cancel);
            this.Controls.Add(this.btn_saveClose);
            this.Controls.Add(this.CorrelScatter);
            this.KeyPreview = true;
            this.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.Name = "CorrelationForm";
            this.Text = "CorrelationForm";
            this.Load += new System.EventHandler(this.CorrelationForm_Load);
            this.Paint += new System.Windows.Forms.PaintEventHandler(this.CorrelationForm_Paint);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.CorrelationForm_KeyDown);
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
        private System.Windows.Forms.Button btn_LaunchHelper;
        private System.Windows.Forms.Button btn_LaunchDrawCorrelation;
    }
}