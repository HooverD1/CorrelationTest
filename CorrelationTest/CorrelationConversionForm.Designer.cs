namespace CorrelationTest
{
    partial class CorrelationConversionForm
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
            this.btnConvert = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.checkboxPreserveOffDiagonal = new System.Windows.Forms.CheckBox();
            this.label_Heading = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // btnConvert
            // 
            this.btnConvert.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnConvert.Location = new System.Drawing.Point(17, 109);
            this.btnConvert.Name = "btnConvert";
            this.btnConvert.Size = new System.Drawing.Size(122, 42);
            this.btnConvert.TabIndex = 0;
            this.btnConvert.Text = "Convert";
            this.btnConvert.UseVisualStyleBackColor = true;
            this.btnConvert.Click += new System.EventHandler(this.btnConvert_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.Location = new System.Drawing.Point(173, 109);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(122, 42);
            this.btnCancel.TabIndex = 1;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // checkboxPreserveOffDiagonal
            // 
            this.checkboxPreserveOffDiagonal.AutoSize = true;
            this.checkboxPreserveOffDiagonal.Location = new System.Drawing.Point(17, 54);
            this.checkboxPreserveOffDiagonal.Name = "checkboxPreserveOffDiagonal";
            this.checkboxPreserveOffDiagonal.Size = new System.Drawing.Size(210, 23);
            this.checkboxPreserveOffDiagonal.TabIndex = 2;
            this.checkboxPreserveOffDiagonal.Text = "Preserve Off-Diagonal Values";
            this.checkboxPreserveOffDiagonal.UseVisualStyleBackColor = true;
            // 
            // label_Heading
            // 
            this.label_Heading.AutoSize = true;
            this.label_Heading.Font = new System.Drawing.Font("Leelawadee UI", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label_Heading.Location = new System.Drawing.Point(12, 9);
            this.label_Heading.Name = "label_Heading";
            this.label_Heading.Size = new System.Drawing.Size(283, 28);
            this.label_Heading.TabIndex = 3;
            this.label_Heading.Text = "Convert to _______ Specification";
            // 
            // CorrelationConversionForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(307, 167);
            this.Controls.Add(this.label_Heading);
            this.Controls.Add(this.checkboxPreserveOffDiagonal);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnConvert);
            this.Font = new System.Drawing.Font("Leelawadee UI", 7.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Name = "CorrelationConversionForm";
            this.Text = "Convert Specification";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnConvert;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.CheckBox checkboxPreserveOffDiagonal;
        private System.Windows.Forms.Label label_Heading;
    }
}