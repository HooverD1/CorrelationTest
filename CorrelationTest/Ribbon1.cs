using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
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
            Dictionary<string, object> sheetData = new Dictionary<string, object>() { { "xlSheet", xlSheet } };
            ICostSheet wbs_sheet = CostSheetFactory.Construct(Sheets.Sheet.GetSheetType(xlSheet), sheetData);
            wbs_sheet.BuildCorrelations();
        }

        private void ExpandCorrel_Click(object sender, RibbonControlEventArgs e)
        {
            Data.CorrelationString.ExpandCorrel(ThisAddIn.MyApp.Selection);
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
            xlSheet.Cells[5, 1] = 1;
            xlSheet.Cells[5, 2] = "E";
            xlSheet.Cells[5, 3] = "Est1";
            xlSheet.Cells[5, 5] = "Normal";
            xlSheet.Cells[5, 6] = 0;
            xlSheet.Cells[5, 7] = 1;
            xlSheet.Cells[7, 1] = 2;
            xlSheet.Cells[7, 2] = "E";
            xlSheet.Cells[7, 3] = "Est2";
            xlSheet.Cells[7, 5] = "Triangular";
            xlSheet.Cells[7, 6] = 10;
            xlSheet.Cells[7, 7] = 30;
            xlSheet.Cells[7, 8] = 20;
            xlSheet.Cells[9, 1] = 2;
            xlSheet.Cells[9, 2] = "E";
            xlSheet.Cells[9, 3] = "Est3";
            xlSheet.Cells[9, 5] = "Normal";
            xlSheet.Cells[9, 6] = 0;
            xlSheet.Cells[9, 7] = 1;
            xlSheet.Cells[11, 1] = 1;
            xlSheet.Cells[11, 2] = "E";
            xlSheet.Cells[11, 3] = "Est4";
            xlSheet.Cells[11, 5] = "Normal";
            xlSheet.Cells[11, 6] = 0;
            xlSheet.Cells[11, 7] = 1;
            xlSheet.Cells[13, 1] = 2;
            xlSheet.Cells[13, 2] = "E";
            xlSheet.Cells[13, 3] = "Est5";
            xlSheet.Cells[13, 5] = "Normal";
            xlSheet.Cells[13, 6] = 0;
            xlSheet.Cells[13, 7] = 1;
            xlSheet.Cells[14, 1] = 2;
            xlSheet.Cells[14, 2] = "E";
            xlSheet.Cells[14, 3] = "Est6";
            xlSheet.Cells[14, 5] = "Lognormal";
            xlSheet.Cells[14, 6] = 0;
            xlSheet.Cells[14, 7] = 1;
            xlSheet.Cells[15, 1] = 3;
            xlSheet.Cells[15, 2] = "E";
            xlSheet.Cells[15, 3] = "Est7";
            xlSheet.Cells[15, 5] = "Normal";
            xlSheet.Cells[15, 6] = 0;
            xlSheet.Cells[15, 7] = 1;
            xlSheet.Cells[16, 1] = 3;
            xlSheet.Cells[16, 2] = "E";
            xlSheet.Cells[16, 3] = "Est8";
            xlSheet.Cells[16, 5] = "Normal";
            xlSheet.Cells[16, 6] = 0;
            xlSheet.Cells[16, 7] = 1;
        }

        private void btnVisualize_Click(object sender, RibbonControlEventArgs e)
        {
            Sheets.CorrelationSheet correlSheet = Sheets.CorrelationSheet.BuildFromExisting();
            correlSheet.VisualizeCorrel();
        }
    }
}
