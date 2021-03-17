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
        public CorrelationConversionForm()
        {
            InitializeComponent();
        }

        private void btnConvert_Click(object sender, EventArgs e)
        {
            //sender is the button itself!

            /*
             * If I want to convert, I need to build the original first, pull its components, and send them to the new one's constructor.
             */

            //Build the  viewmodel for the existing correlation sheet.
            //try
            //{
                var correlSheet = Sheets.CorrelationSheet.ConstructFromXlCorrelationSheet();
                correlSheet.ConvertCorrelation();
                this.Close();
            //}
            //catch(Exception except)
            //{
            //    if (MyGlobals.DebugMode)
            //        throw except;
            //}
        }
    }
}
