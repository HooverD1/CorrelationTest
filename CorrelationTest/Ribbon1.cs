using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace CorrelationTest
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void BuildCorrelation_Click(object sender, RibbonControlEventArgs e)
        {
          
        }

        private void ExpandCorrel_Click(object sender, RibbonControlEventArgs e)
        {
                SendKeys.Send("{ESC}");
                //Need correlation string to expand depending on the value in Selection
                Excel.Range selection = ThisAddIn.MyApp.Selection;
                Data.CorrelationString cs = Data.CorrelationString.ConstructFromExisting(Convert.ToString(selection.Value));
                cs.Expand(selection);
        }

        private void CollapseCorrel_Click(object sender, RibbonControlEventArgs e)
        {
            //cancel edits
            Sheets.CorrelationSheet.CollapseToSheet();
        }

        private void FakeFields_Click(object sender, RibbonControlEventArgs e)
        {
            //Search for existing EST_1 sheet
            Excel.Worksheet est_1 = ExtensionMethods.GetWorksheet("EST_1", SheetType.Estimate);
            Excel.Worksheet wbs_1 = ExtensionMethods.GetWorksheet("WBS_1", SheetType.WBS);

            est_1.Cells[4, 1] = "ID";
            est_1.Cells[4, 4] = "Name";
            est_1.Cells[4, 7] = "Distribution";
            est_1.Cells[4, 8] = "Param1";
            est_1.Cells[4, 9] = "Param2";
            est_1.Cells[4, 10] = "Param3";

            est_1.Cells[5, 1] = $"E|{DateTime.Now.ToUniversalTime().ToString("ddMMyy")}{DateTime.Now.ToUniversalTime().ToString("HH:mm:ss.fff")}";
            est_1.Cells[5, 2] = 4;  //# of inputs
            est_1.Cells[5, 3] = "CE";
            est_1.Cells[5, 4] = "Est1";
            est_1.Cells[5, 7] = "Normal";
            est_1.Cells[5, 8] = 0;
            est_1.Cells[5, 9] = 1;
            for (int k = 14; k < 20; k++)
            {
                est_1.Cells[5, k] = 7;
            }
            System.Threading.Thread.Sleep(1);
            est_1.Cells[6, 1] = $"E|{DateTime.Now.ToUniversalTime().ToString("ddMMyy")}{DateTime.Now.ToUniversalTime().ToString("HH:mm:ss.fff")}";
            est_1.Cells[6, 3] = "I";
            est_1.Cells[6, 4] = "Est2";
            est_1.Cells[6, 7] = "Triangular";
            est_1.Cells[6, 8] = 10;
            est_1.Cells[6, 9] = 30;
            est_1.Cells[6, 10] = 20;
            for (int k = 14; k < 20; k++)
            {
                est_1.Cells[6, k] = 7;
            }
            System.Threading.Thread.Sleep(1);
            est_1.Cells[7, 1] = $"E|{DateTime.Now.ToUniversalTime().ToString("ddMMyy")}{DateTime.Now.ToUniversalTime().ToString("HH:mm:ss.fff")}";
            est_1.Cells[7, 3] = "I";
            est_1.Cells[7, 4] = "Est3";
            est_1.Cells[7, 7] = "Triangular";
            est_1.Cells[7, 8] = 10;
            est_1.Cells[7, 9] = 30;
            est_1.Cells[7, 10] = 20;
            for (int k = 14; k < 20; k++)
            {
                est_1.Cells[7, k] = 7;
            }
            System.Threading.Thread.Sleep(1);
            est_1.Cells[8, 1] = $"E|{DateTime.Now.ToUniversalTime().ToString("ddMMyy")}{DateTime.Now.ToUniversalTime().ToString("HH:mm:ss.fff")}";
            est_1.Cells[8, 3] = "I";
            est_1.Cells[8, 4] = "Est4";
            est_1.Cells[8, 7] = "Triangular";
            est_1.Cells[8, 8] = 10;
            est_1.Cells[8, 9] = 30;
            est_1.Cells[8, 10] = 20;
            for (int k = 14; k < 20; k++)
            {
                est_1.Cells[8, k] = 7;
            }
            System.Threading.Thread.Sleep(1);
            est_1.Cells[9, 1] = $"E|{DateTime.Now.ToUniversalTime().ToString("ddMMyy")}{DateTime.Now.ToUniversalTime().ToString("HH:mm:ss.fff")}";
            est_1.Cells[9, 3] = "I";
            est_1.Cells[9, 4] = "Est5";
            est_1.Cells[9, 7] = "Normal";
            est_1.Cells[9, 8] = 0;
            est_1.Cells[9, 9] = 1;
            for (int k = 14; k < 20; k++)
            {
                est_1.Cells[9, k] = 7;
            }
            System.Threading.Thread.Sleep(1);
            est_1.Cells[10, 1] = $"E|{DateTime.Now.ToUniversalTime().ToString("ddMMyy")}{DateTime.Now.ToUniversalTime().ToString("HH:mm:ss.fff")}";
            est_1.Cells[10, 2] = 4;
            est_1.Cells[10, 3] = "CE";
            est_1.Cells[10, 4] = "Est5.2";
            est_1.Cells[10, 7] = "Normal";
            est_1.Cells[10, 8] = 0;
            est_1.Cells[10, 9] = 1;
            for (int k = 14; k < 20; k++)
            {
                est_1.Cells[10, k] = 7;
            }
            System.Threading.Thread.Sleep(1);
            est_1.Cells[11, 1] = $"E|{DateTime.Now.ToUniversalTime().ToString("ddMMyy")}{DateTime.Now.ToUniversalTime().ToString("HH:mm:ss.fff")}";
            est_1.Cells[11, 3] = "I";
            est_1.Cells[11, 4] = "Est6";
            est_1.Cells[11, 7] = "Normal";
            est_1.Cells[11, 8] = 0;
            est_1.Cells[11, 9] = 1;
            for (int k = 14; k < 20; k++)
            {
                est_1.Cells[11, k] = 7;
            }
            System.Threading.Thread.Sleep(1);
            est_1.Cells[12, 1] = $"E|{DateTime.Now.ToUniversalTime().ToString("ddMMyy")}{DateTime.Now.ToUniversalTime().ToString("HH:mm:ss.fff")}";
            est_1.Cells[12, 3] = "I";
            est_1.Cells[12, 4] = "Est7";
            est_1.Cells[12, 7] = "Normal";
            est_1.Cells[12, 8] = 0;
            est_1.Cells[12, 9] = 1;
            for (int k = 14; k < 20; k++)
            {
                est_1.Cells[12, k] = 7;
            }
            System.Threading.Thread.Sleep(1);
            est_1.Cells[13, 1] = $"E|{DateTime.Now.ToUniversalTime().ToString("ddMMyy")}{DateTime.Now.ToUniversalTime().ToString("HH:mm:ss.fff")}";
            est_1.Cells[13, 3] = "I";
            est_1.Cells[13, 4] = "Est8";
            est_1.Cells[13, 7] = "Normal";
            est_1.Cells[13, 8] = 0;
            est_1.Cells[13, 9] = 1;
            for (int k = 14; k < 20; k++)
            {
                est_1.Cells[13, k] = 7;
            }
            System.Threading.Thread.Sleep(1);
            est_1.Cells[14, 1] = $"E|{DateTime.Now.ToUniversalTime().ToString("ddMMyy")}{DateTime.Now.ToUniversalTime().ToString("HH:mm:ss.fff")}";
            est_1.Cells[14, 3] = "I";
            est_1.Cells[14, 4] = "Est9";
            est_1.Cells[14, 7] = "Lognormal";
            est_1.Cells[14, 8] = 0;
            est_1.Cells[14, 9] = 1;
            for (int k = 14; k < 20; k++)
            {
                est_1.Cells[14, k] = 7;
            }
            System.Threading.Thread.Sleep(1);
            est_1.Cells[15, 1] = $"CE|{DateTime.Now.ToUniversalTime().ToString("ddMMyy")}{ DateTime.Now.ToUniversalTime().ToString("HH:mm:ss.fff")}";
            est_1.Cells[15, 2] = 1;
            est_1.Cells[15, 3] = "CE";
            est_1.Cells[15, 4] = "Est10";
            est_1.Cells[15, 7] = "Normal";
            est_1.Cells[15, 8] = 0;
            est_1.Cells[15, 9] = 1;
            for (int k = 14; k < 20; k++)
            { 
                est_1.Cells[15, k] = 7;
            }

            System.Threading.Thread.Sleep(1);
            est_1.Cells[16, 1] = $"E|{DateTime.Now.ToUniversalTime().ToString("ddMMyy")}{ DateTime.Now.ToUniversalTime().ToString("HH:mm:ss.fff")}";
            est_1.Cells[16, 3] = "I";
            est_1.Cells[16, 4] = "Est11";
            est_1.Cells[16, 7] = "Normal";
            est_1.Cells[16, 8] = 0;
            est_1.Cells[16, 9] = 1;
            for (int k = 14; k < 20; k++)
            {
                est_1.Cells[16, k] = 7;
            }

            est_1.Activate();

            wbs_1.Cells[4, 1] = "ID";
            wbs_1.Cells[4, 2] = "Level";
            wbs_1.Cells[4, 4] = "Name";
            wbs_1.Cells[4, 7] = "Distribution";
            wbs_1.Cells[4, 8] = "Param1";
            wbs_1.Cells[4, 9] = "Param2";
            wbs_1.Cells[4, 10] = "Param3";

            System.Threading.Thread.Sleep(1);
            wbs_1.Cells[5, 1] = $"S|{DateTime.Now.ToUniversalTime().ToString("ddMMyy")}{DateTime.Now.ToUniversalTime().ToString("HH: mm:ss.fff")}";
            wbs_1.Cells[5, 2] = 1;
            wbs_1.Cells[5, 3] = "S";

            System.Threading.Thread.Sleep(1);
            wbs_1.Cells[6, 1] = $"W|{DateTime.Now.ToUniversalTime().ToString("ddMMyy")}{DateTime.Now.ToUniversalTime().ToString("HH: mm:ss.fff")}";
            wbs_1.Cells[6, 2] = 2;
            wbs_1.Cells[6, 3] = "CE";
            wbs_1.Cells[6, 4] = "Est2";
            wbs_1.Cells[6, 7] = "Triangular";
            wbs_1.Cells[6, 8] = 10;
            wbs_1.Cells[6, 9] = 30;
            wbs_1.Cells[6, 10] = 20;
            for (int k = 14; k < 20; k++)
            {
                wbs_1.Cells[6, k] = 7;
            }
            System.Threading.Thread.Sleep(1);
            wbs_1.Cells[7, 1] = $"W|{DateTime.Now.ToUniversalTime().ToString("ddMMyy")}{DateTime.Now.ToUniversalTime().ToString("HH: mm:ss.fff")}";
            wbs_1.Cells[7, 2] = 2;
            wbs_1.Cells[7, 3] = "CE";
            wbs_1.Cells[7, 4] = "Est3";
            wbs_1.Cells[7, 7] = "Triangular";
            wbs_1.Cells[7, 8] = 10;
            wbs_1.Cells[7, 9] = 30;
            wbs_1.Cells[7, 10] = 20;
            for (int k = 14; k < 20; k++)
            {
                wbs_1.Cells[7, k] = 7;
            }
            System.Threading.Thread.Sleep(1);
            wbs_1.Cells[8, 1] = $"W|{DateTime.Now.ToUniversalTime().ToString("ddMMyy")}{DateTime.Now.ToUniversalTime().ToString("HH: mm:ss.fff")}";
            wbs_1.Cells[8, 2] = 2;
            wbs_1.Cells[8, 3] = "CE";
            wbs_1.Cells[8, 4] = "Est4";
            wbs_1.Cells[8, 7] = "Triangular";
            wbs_1.Cells[8, 8] = 10;
            wbs_1.Cells[8, 9] = 30;
            wbs_1.Cells[8, 10] = 20;
            for (int k = 14; k < 20; k++)
            {
                wbs_1.Cells[8, k] = 7;
            }
            System.Threading.Thread.Sleep(1);
            wbs_1.Cells[9, 1] = $"W|{DateTime.Now.ToUniversalTime().ToString("ddMMyy")}{DateTime.Now.ToUniversalTime().ToString("HH: mm:ss.fff")}";
            wbs_1.Cells[9, 2] = 2;
            wbs_1.Cells[9, 3] = "CE";
            wbs_1.Cells[9, 4] = "Est5";
            wbs_1.Cells[9, 7] = "Normal";
            wbs_1.Cells[9, 8] = 0;
            wbs_1.Cells[9, 9] = 1;
            for (int k = 14; k < 20; k++)
            {
                wbs_1.Cells[9, k] = 7;
            }

            System.Threading.Thread.Sleep(1);
            wbs_1.Cells[10, 1] = $"S|{DateTime.Now.ToUniversalTime().ToString("ddMMyy")}{DateTime.Now.ToUniversalTime().ToString("HH: mm:ss.fff")}";
            wbs_1.Cells[10, 2] = 1;
            wbs_1.Cells[10, 3] = "S";

            System.Threading.Thread.Sleep(1);
            wbs_1.Cells[11, 1] = $"W|{DateTime.Now.ToUniversalTime().ToString("ddMMyy")}{DateTime.Now.ToUniversalTime().ToString("HH: mm:ss.fff")}";
            wbs_1.Cells[11, 2] = 2;
            wbs_1.Cells[11, 3] = "CE";
            wbs_1.Cells[11, 4] = "Est6";
            wbs_1.Cells[11, 7] = "Normal";
            wbs_1.Cells[11, 8] = 0;
            wbs_1.Cells[11, 9] = 1;
            for (int k = 14; k < 20; k++)
            {
                wbs_1.Cells[11, k] = 7;
            }
            System.Threading.Thread.Sleep(1);
            wbs_1.Cells[12, 1] = $"W|{DateTime.Now.ToUniversalTime().ToString("ddMMyy")}{DateTime.Now.ToUniversalTime().ToString("HH: mm:ss.fff")}";
            wbs_1.Cells[12, 2] = 2;
            wbs_1.Cells[12, 3] = "CE";
            wbs_1.Cells[12, 4] = "Est7";
            wbs_1.Cells[12, 7] = "Normal";
            wbs_1.Cells[12, 8] = 0;
            wbs_1.Cells[12, 9] = 1;
            for (int k = 14; k < 20; k++)
            {
                wbs_1.Cells[12, k] = 7;
            }
            System.Threading.Thread.Sleep(1);
            wbs_1.Cells[13, 1] = $"W|{DateTime.Now.ToUniversalTime().ToString("ddMMyy")}{DateTime.Now.ToUniversalTime().ToString("HH: mm:ss.fff")}";
            wbs_1.Cells[13, 2] = 2;
            wbs_1.Cells[13, 3] = "CE";
            wbs_1.Cells[13, 4] = "Est8";
            wbs_1.Cells[13, 7] = "Normal";
            wbs_1.Cells[13, 8] = 0;
            wbs_1.Cells[13, 9] = 1;
            for (int k = 14; k < 20; k++)
            {
                wbs_1.Cells[13, k] = 7;
            }
            System.Threading.Thread.Sleep(1);
            wbs_1.Cells[14, 1] = $"W|{DateTime.Now.ToUniversalTime().ToString("ddMMyy")}{DateTime.Now.ToUniversalTime().ToString("HH: mm:ss.fff")}";
            wbs_1.Cells[14, 2] = 2;
            wbs_1.Cells[14, 3] = "CE";
            wbs_1.Cells[14, 4] = "Est9";
            wbs_1.Cells[14, 7] = "Lognormal";
            wbs_1.Cells[14, 8] = 0;
            wbs_1.Cells[14, 9] = 1;
            for (int k = 14; k < 20; k++)
            {
                wbs_1.Cells[14, k] = 7;
            }

            System.Threading.Thread.Sleep(1);
            wbs_1.Cells[15, 1] = $"S|{DateTime.Now.ToUniversalTime().ToString("ddMMyy")}{DateTime.Now.ToUniversalTime().ToString("HH: mm:ss.fff")}";
            wbs_1.Cells[15, 2] = 1;
            wbs_1.Cells[15, 3] = "S";

            System.Threading.Thread.Sleep(1);
            wbs_1.Cells[16, 1] = $"W|{DateTime.Now.ToUniversalTime().ToString("ddMMyy")}{DateTime.Now.ToUniversalTime().ToString("HH: mm:ss.fff")}";
            wbs_1.Cells[16, 2] = 2;
            wbs_1.Cells[16, 3] = "CE";
            wbs_1.Cells[16, 4] = "Est11";
            wbs_1.Cells[16, 7] = "Normal";
            wbs_1.Cells[16, 8] = 0;
            wbs_1.Cells[16, 9] = 1;
            for (int k = 14; k < 20; k++)
            {
                wbs_1.Cells[16, k] = 7;
            }

            //Goal: Build the correlation strings on each example sheet
            //Steps
            //1 -- Build the sheet object -- est_1 is the xlSheet; construct the sheet object from it
            CostSheet estimateSheet_example = CostSheet.Construct(est_1);
            //2 -- Manually load the estimate objects to the sheet object, including their SubEstimates
            estimateSheet_example.PrintDefaultCorrelStrings();
            //3 -- Build default CorrelStrings for estimates attached to the sheet object


            //Repeat for wbs_1
            CostSheet wbsSheet_example = CostSheet.Construct(wbs_1);
            wbsSheet_example.PrintDefaultCorrelStrings();

        }

        private void btnVisualize_Click(object sender, RibbonControlEventArgs e)
        {
            Sheets.CorrelationSheet correlSheet = Sheets.CorrelationSheet.Construct();
            if (correlSheet == null)
                return;

            correlSheet.VisualizeCorrel();
        }

    }
}
