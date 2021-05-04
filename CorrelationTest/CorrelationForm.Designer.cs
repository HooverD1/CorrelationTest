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
            System.Windows.Forms.DataVisualization.Charting.ChartArea chartArea3 = new System.Windows.Forms.DataVisualization.Charting.ChartArea();
            System.Windows.Forms.DataVisualization.Charting.Legend legend3 = new System.Windows.Forms.DataVisualization.Charting.Legend();
            System.Windows.Forms.DataVisualization.Charting.Series series3 = new System.Windows.Forms.DataVisualization.Charting.Series();
            this.CorrelScatter = new System.Windows.Forms.DataVisualization.Charting.Chart();
            this.btn_saveClose = new System.Windows.Forms.Button();
            this.btn_Cancel = new System.Windows.Forms.Button();
            this.numericUpDown_CorrelValue = new System.Windows.Forms.NumericUpDown();
            this.groupBox_CorrelCoef = new System.Windows.Forms.GroupBox();
            ((System.ComponentModel.ISupportInitialize)(this.CorrelScatter)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDown_CorrelValue)).BeginInit();
            this.groupBox_CorrelCoef.SuspendLayout();
            this.SuspendLayout();
            // 
            // CorrelScatter
            // 
            chartArea3.Name = "ChartArea1";
            this.CorrelScatter.ChartAreas.Add(chartArea3);
            legend3.Name = "Legend1";
            this.CorrelScatter.Legends.Add(legend3);
            this.CorrelScatter.Location = new System.Drawing.Point(12, 12);
            this.CorrelScatter.Name = "CorrelScatter";
            series3.ChartArea = "ChartArea1";
            series3.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Point;
            series3.Legend = "Legend1";
            series3.Name = "CorrelSeries";
            series3.XValueType = System.Windows.Forms.DataVisualization.Charting.ChartValueType.Double;
            series3.YValueType = System.Windows.Forms.DataVisualization.Charting.ChartValueType.Double;
            this.CorrelScatter.Series.Add(series3);
            this.CorrelScatter.Size = new System.Drawing.Size(1075, 642);
            this.CorrelScatter.TabIndex = 0;
            this.CorrelScatter.Text = "Scatterplot";
            this.CorrelScatter.Click += new System.EventHandler(this.CorrelScatter_Click);
            // 
            // btn_saveClose
            // 
            this.btn_saveClose.Location = new System.Drawing.Point(887, 509);
            this.btn_saveClose.Name = "btn_saveClose";
            this.btn_saveClose.Size = new System.Drawing.Size(184, 38);
            this.btn_saveClose.TabIndex = 1;
            this.btn_saveClose.Text = "Save";
            this.btn_saveClose.UseVisualStyleBackColor = true;
            this.btn_saveClose.Click += new System.EventHandler(this.btn_saveClose_Click);
            // 
            // btn_Cancel
            // 
            this.btn_Cancel.Location = new System.Drawing.Point(887, 553);
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
            this.numericUpDown_CorrelValue.Name = "numericUpDown_CorrelValue";
            this.numericUpDown_CorrelValue.Size = new System.Drawing.Size(143, 34);
            this.numericUpDown_CorrelValue.TabIndex = 3;
            this.numericUpDown_CorrelValue.ValueChanged += new System.EventHandler(this.numericUpDown_CorrelValue_ValueChanged);
            this.numericUpDown_CorrelValue.MouseDown += new System.Windows.Forms.MouseEventHandler(this.numericUpDown_CorrelValue_MouseDown);
            // 
            // groupBox_CorrelCoef
            // 
            this.groupBox_CorrelCoef.BackColor = System.Drawing.SystemColors.Info;
            this.groupBox_CorrelCoef.Controls.Add(this.numericUpDown_CorrelValue);
            this.groupBox_CorrelCoef.Location = new System.Drawing.Point(887, 397);
            this.groupBox_CorrelCoef.Name = "groupBox_CorrelCoef";
            this.groupBox_CorrelCoef.Size = new System.Drawing.Size(184, 106);
            this.groupBox_CorrelCoef.TabIndex = 5;
            this.groupBox_CorrelCoef.TabStop = false;
            this.groupBox_CorrelCoef.Text = "Correlation Coefficient";
            this.groupBox_CorrelCoef.Enter += new System.EventHandler(this.groupBox_CorrelCoef_Enter);
            // 
            // CorrelationForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.Window;
            this.ClientSize = new System.Drawing.Size(1083, 660);
            this.Controls.Add(this.groupBox_CorrelCoef);
            this.Controls.Add(this.btn_Cancel);
            this.Controls.Add(this.btn_saveClose);
            this.Controls.Add(this.CorrelScatter);
            this.Name = "CorrelationForm";
            this.Text = "CorrelationForm";
            this.Load += new System.EventHandler(this.CorrelationForm_Load);
            ((System.ComponentModel.ISupportInitialize)(this.CorrelScatter)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDown_CorrelValue)).EndInit();
            this.groupBox_CorrelCoef.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.DataVisualization.Charting.Chart CorrelScatter;
        private System.Windows.Forms.Button btn_saveClose;
        private System.Windows.Forms.Button btn_Cancel;
        private System.Windows.Forms.NumericUpDown numericUpDown_CorrelValue;
        private System.Windows.Forms.GroupBox groupBox_CorrelCoef;
    }
}