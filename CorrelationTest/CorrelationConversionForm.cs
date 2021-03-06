using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace CorrelationTest
{
    public partial class CorrelationConversionForm : Form
    {
        public CorrelationConversionForm(Sheets.CorrelationSheet containingSheet)
        {
            InitializeComponent();
            SetupHeaderLabel(containingSheet);
            SetupActiveOptions(containingSheet);
        }

        private void SetupHeaderLabel(Sheets.CorrelationSheet containingSheet)
        {
            switch (containingSheet)
            {
                case Sheets.CorrelationSheet_DP dp:
                    this.label_Heading.Text = "Convert to Matrix Specification";
                    break;
                case Sheets.CorrelationSheet_DM dm:
                    this.label_Heading.Text = "Convert to Pairwise Specification";
                    break;
                case Sheets.CorrelationSheet_CP cp:
                    this.label_Heading.Text = "Convert to Matrix Specification";
                    break;
                case Sheets.CorrelationSheet_CM cm:
                    this.label_Heading.Text = "Convert to Pairwise Specification";
                    break;
            }
        }

        private void SetupActiveOptions(Sheets.CorrelationSheet containingSheet)
        {
            switch (containingSheet)
            {
                case Sheets.CorrelationSheet_DP dp:
                    checkboxPreserveOffDiagonal.Enabled = false;
                    break;
                case Sheets.CorrelationSheet_DM dm:
                    checkboxPreserveOffDiagonal.Enabled = true;
                    break;
                case Sheets.CorrelationSheet_CP cp:
                    checkboxPreserveOffDiagonal.Enabled = false;
                    break;
                case Sheets.CorrelationSheet_CM cm:
                    checkboxPreserveOffDiagonal.Enabled = true;
                    break;
            }
        }

        private void btnConvert_Click(object sender, EventArgs e)
        {
            ExtensionMethods.TurnOffUpdating();
            var correlSheet = Sheets.CorrelationSheet.ConstructFromXlCorrelationSheet();
            correlSheet.ConvertCorrelation(this.checkboxPreserveOffDiagonal.Checked);
            this.Close();
            ExtensionMethods.TurnOnUpdating();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
