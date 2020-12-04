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
            //build all default correlations if they don't exist
            Excel.Worksheet xlSheet = ThisAddIn.MyApp.ActiveSheet;
            SheetType sheetType = Sheets.Sheet.GetSheetType(xlSheet);
            if(sheetType == SheetType.WBS)
            {
                Dictionary<string, object> sheetData = new Dictionary<string, object>() { { "SheetType", sheetType }, { "xlSheet", xlSheet } };
                ICostSheet wbs_sheet = CostSheetFactory.Construct(sheetData);
                wbs_sheet.BuildCorrelations();
            }
            else if(sheetType == SheetType.Estimate)
            {

            }
            else
            {

            }
        }

        private void ExpandCorrel_Click(object sender, RibbonControlEventArgs e)
        {
            SendKeys.Send("{ESC}");
            Data.CorrelationString_Inputs.ExpandCorrel(ThisAddIn.MyApp.Selection);
        }

        private void CollapseCorrel_Click(object sender, RibbonControlEventArgs e)
        {
            //cancel edits
            Sheets.CorrelationSheet.CollapseToSheet();
        }

        private void FakeFields_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Worksheet xlSheet = ThisAddIn.MyApp.ActiveSheet;
            xlSheet.Cells[1, 1] = "$WBS";
            xlSheet.Cells[5, 2] = 1;
            xlSheet.Cells[5, 3] = "E";
            xlSheet.Cells[5, 4] = "Est1";
            xlSheet.Cells[5, 6] = "Normal";
            xlSheet.Cells[5, 7] = 0;
            xlSheet.Cells[5, 8] = 1;
            for (int i = 12; i < 22; i++)
                xlSheet.Cells[5, i].value = 7;
            xlSheet.Cells[7, 2] = 2;
            xlSheet.Cells[7, 3] = "E";
            xlSheet.Cells[7, 4] = "Est2";
            xlSheet.Cells[7, 6] = "Triangular";
            xlSheet.Cells[7, 7] = 10;
            xlSheet.Cells[7, 8] = 30;
            xlSheet.Cells[7, 9] = 20;
            for (int i = 12; i < 22; i++)
                xlSheet.Cells[7, i].value = 7;
            xlSheet.Cells[9, 2] = 2;
            xlSheet.Cells[9, 3] = "E";
            xlSheet.Cells[9, 4] = "Est3";
            xlSheet.Cells[9, 6] = "Normal";
            xlSheet.Cells[9, 7] = 0;
            xlSheet.Cells[9, 8] = 1;
            for (int i = 12; i < 22; i++)
                xlSheet.Cells[9, i].value = 7;
            xlSheet.Cells[11, 2] = 1;
            xlSheet.Cells[11, 3] = "E";
            xlSheet.Cells[11, 4] = "Est4";
            xlSheet.Cells[11, 6] = "Normal";
            xlSheet.Cells[11, 7] = 0;
            xlSheet.Cells[11, 8] = 1;
            for (int i = 12; i < 22; i++)
                xlSheet.Cells[11, i].value = 7;
            xlSheet.Cells[13, 2] = 2;
            xlSheet.Cells[13, 3] = "E";
            xlSheet.Cells[13, 4] = "Est5";
            xlSheet.Cells[13, 6] = "Normal";
            xlSheet.Cells[13, 7] = 0;
            xlSheet.Cells[13, 8] = 1;
            for (int i = 12; i < 22; i++)
                xlSheet.Cells[13, i].value = 7;
            xlSheet.Cells[14, 2] = 2;
            xlSheet.Cells[14, 3] = "E";
            xlSheet.Cells[14, 4] = "Est6";
            xlSheet.Cells[14, 6] = "Lognormal";
            xlSheet.Cells[14, 7] = 0;
            xlSheet.Cells[14, 8] = 1;
            for (int i = 12; i < 22; i++)
                xlSheet.Cells[14, i].value = 7;
            xlSheet.Cells[15, 2] = 3;
            xlSheet.Cells[15, 3] = "E";
            xlSheet.Cells[15, 4] = "Est7";
            xlSheet.Cells[15, 6] = "Normal";
            xlSheet.Cells[15, 7] = 0;
            xlSheet.Cells[15, 8] = 1;
            for (int i = 12; i < 22; i++)
                xlSheet.Cells[15, i].value = 7;
            xlSheet.Cells[16, 2] = 3;
            xlSheet.Cells[16, 3] = "E";
            xlSheet.Cells[16, 4] = "Est8";
            xlSheet.Cells[16, 6] = "Normal";
            xlSheet.Cells[16, 7] = 0;
            xlSheet.Cells[16, 8] = 1;
            for (int i = 12; i < 22; i++)
                xlSheet.Cells[16, i].value = 7;
        }

        private void btnVisualize_Click(object sender, RibbonControlEventArgs e)
        {
            Sheets.CorrelationSheet correlSheet = Sheets.CorrelationSheet.BuildFromExisting();
            if (correlSheet == null)
                return;

            correlSheet.VisualizeCorrel();
        }
    }
}
