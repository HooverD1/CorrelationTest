﻿using System;
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
            
            Excel.Range selection = ThisAddIn.MyApp.Selection;
            SheetType sheetType = ExtensionMethods.GetSheetType(selection.Worksheet);
            DisplayCoords dispCoords = DisplayCoords.ConstructDisplayCoords(sheetType);
            CostSheet sheetObj = CostSheet.ConstructFromXlCostSheet(selection.Worksheet);
            
            Item selectedItem = (from Item item in sheetObj.Items where item.xlRow.Row == selection.Row select item).First();
            CorrelationType correlType;
            if(!(selectedItem is ISub))
            {
                //
                return;
            }
            if (((ISub)selectedItem).Parent is IHasCostCorrelations)
            {
                correlType = CorrelationType.Cost;
            }
            else if (((ISub)selectedItem).Parent is IHasDurationCorrelations)
            {
                correlType = CorrelationType.Duration;
            }
            else if (selection.Column == dispCoords.PhasingCorrel_Offset && ((ISub)selectedItem) is IHasPhasingCorrelations)
            {
                correlType = CorrelationType.Phasing;
            }
            else
            {
                correlType = CorrelationType.Null;
                throw new Exception("Unknown Correlation Type");
            }

            switch (correlType)
            {
                case CorrelationType.Cost:
                    ((ISub)selectedItem).Parent.Expand(correlType);
                    break;
                case CorrelationType.Duration:
                    ((ISub)selectedItem).Parent.Expand(correlType);
                    break;
                case CorrelationType.Phasing:
                    ((IHasPhasingCorrelations)selectedItem).Expand(correlType);
                    break;
                case CorrelationType.Null:      //Not selecting a correlation column
                    return;     
                default:
                    throw new Exception("Unknown correlation expand issue");
            }
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
            DisplayCoords edc = DisplayCoords.ConstructDisplayCoords(SheetType.Estimate);
            Excel.Worksheet wbs_1 = ExtensionMethods.GetWorksheet("WBS_1", SheetType.WBS);
            DisplayCoords wdc = DisplayCoords.ConstructDisplayCoords(SheetType.WBS);
            
            est_1.Cells[4, edc.ID_Offset] = "ID";
            est_1.Cells[4, edc.Name_Offset] = "Name";
            est_1.Cells[4, edc.Distribution_Offset] = "Distribution";
            est_1.Cells[4, edc.Distribution_Offset+1] = "Param1";
            est_1.Cells[4, edc.Distribution_Offset+2] = "Param2";
            est_1.Cells[4, edc.Distribution_Offset+3] = "Param3";

            est_1.Cells[5, edc.ID_Offset] = $"DH|E|{ThisAddIn.MyApp.UserName}|{DateTime.Now.ToUniversalTime().ToString("ddMMyy")}{DateTime.Now.ToUniversalTime().ToString("HH:mm:ss.fff")}";
            est_1.Cells[5, edc.Type_Offset] = "SACE";
            est_1.Cells[5, edc.Name_Offset] = "Est1";
            est_1.Cells[5, edc.Level_Offset] = 4;
            est_1.Cells[5, edc.Distribution_Offset] = "Normal";
            est_1.Cells[5, edc.Distribution_Offset + 1] = 0;
            est_1.Cells[5, edc.Distribution_Offset + 2] = 1;

            System.Threading.Thread.Sleep(1);
            est_1.Cells[6, edc.ID_Offset] = $"DH|E|{ThisAddIn.MyApp.UserName}|{DateTime.Now.ToUniversalTime().ToString("ddMMyy")}{DateTime.Now.ToUniversalTime().ToString("HH:mm:ss.fff")}";
            est_1.Cells[6, edc.Level_Offset] = 3;  //# of inputs
            est_1.Cells[6, edc.Type_Offset] = "CE";
            est_1.Cells[6, edc.Name_Offset] = "Est1";
            est_1.Cells[6, edc.Distribution_Offset] = "Normal";
            est_1.Cells[6, edc.Distribution_Offset + 1] = 0;
            est_1.Cells[6, edc.Distribution_Offset + 2] = 1;

            System.Threading.Thread.Sleep(1);
            est_1.Cells[7, edc.ID_Offset] = $"DH|E|{ThisAddIn.MyApp.UserName}|{DateTime.Now.ToUniversalTime().ToString("ddMMyy")}{DateTime.Now.ToUniversalTime().ToString("HH:mm:ss.fff")}";
            est_1.Cells[7, edc.Type_Offset] = "I";
            est_1.Cells[7, edc.Name_Offset] = "Est3";
            est_1.Cells[7, edc.Distribution_Offset] = "Triangular";
            est_1.Cells[7, edc.Distribution_Offset + 1] = 10;
            est_1.Cells[7, edc.Distribution_Offset + 2] = 30;
            est_1.Cells[7, edc.Distribution_Offset + 3] = 20;

            System.Threading.Thread.Sleep(1);
            est_1.Cells[8, edc.ID_Offset] = $"DH|E|{ThisAddIn.MyApp.UserName}|{DateTime.Now.ToUniversalTime().ToString("ddMMyy")}{DateTime.Now.ToUniversalTime().ToString("HH:mm:ss.fff")}";
            est_1.Cells[8, edc.Type_Offset] = "I";
            est_1.Cells[8, edc.Name_Offset] = "Est4";
            est_1.Cells[8, edc.Distribution_Offset] = "Triangular";
            est_1.Cells[8, edc.Distribution_Offset + 1] = 10;
            est_1.Cells[8, edc.Distribution_Offset + 2] = 30;
            est_1.Cells[8, edc.Distribution_Offset + 3] = 20;

            System.Threading.Thread.Sleep(1);
            est_1.Cells[9, edc.ID_Offset] = $"DH|E|{ThisAddIn.MyApp.UserName}|{DateTime.Now.ToUniversalTime().ToString("ddMMyy")}{DateTime.Now.ToUniversalTime().ToString("HH:mm:ss.fff")}";
            est_1.Cells[9, edc.Type_Offset] = "I";
            est_1.Cells[9, edc.Name_Offset] = "Est5";
            est_1.Cells[9, edc.Distribution_Offset] = "Normal";
            est_1.Cells[9, edc.Distribution_Offset + 1] = 0;
            est_1.Cells[9, edc.Distribution_Offset + 2] = 1;

            System.Threading.Thread.Sleep(1);
            est_1.Cells[10, edc.ID_Offset] = $"DH|E|{ThisAddIn.MyApp.UserName}|{DateTime.Now.ToUniversalTime().ToString("ddMMyy")}{DateTime.Now.ToUniversalTime().ToString("HH:mm:ss.fff")}";
            est_1.Cells[10, edc.Level_Offset] = 4;
            est_1.Cells[10, edc.Type_Offset] = "SE";
            est_1.Cells[10, edc.Name_Offset] = "Est5.2";
            est_1.Cells[10, edc.Distribution_Offset] = "Normal";
            est_1.Cells[10, edc.Distribution_Offset + 1] = 0;
            est_1.Cells[10, edc.Distribution_Offset + 2] = 1;

            System.Threading.Thread.Sleep(1);
            est_1.Cells[11, edc.ID_Offset] = $"DH|E|{ThisAddIn.MyApp.UserName}|{DateTime.Now.ToUniversalTime().ToString("ddMMyy")}{DateTime.Now.ToUniversalTime().ToString("HH:mm:ss.fff")}";
            est_1.Cells[11, edc.Type_Offset] = "I";
            est_1.Cells[11, edc.Name_Offset] = "Est6";
            est_1.Cells[11, edc.Distribution_Offset] = "Normal";
            est_1.Cells[11, edc.Distribution_Offset + 1] = 0;
            est_1.Cells[11, edc.Distribution_Offset + 2] = 1;

            System.Threading.Thread.Sleep(1);
            est_1.Cells[12, edc.ID_Offset] = $"DH|E|{ThisAddIn.MyApp.UserName}|{DateTime.Now.ToUniversalTime().ToString("ddMMyy")}{DateTime.Now.ToUniversalTime().ToString("HH:mm:ss.fff")}";
            est_1.Cells[12, edc.Type_Offset] = "I";
            est_1.Cells[12, edc.Name_Offset] = "Est7";
            est_1.Cells[12, edc.Distribution_Offset] = "Normal";
            est_1.Cells[12, edc.Distribution_Offset + 1] = 0;
            est_1.Cells[12, edc.Distribution_Offset + 2] = 1;

            System.Threading.Thread.Sleep(1);
            est_1.Cells[13, edc.ID_Offset] = $"DH|E|{ThisAddIn.MyApp.UserName}|{DateTime.Now.ToUniversalTime().ToString("ddMMyy")}{DateTime.Now.ToUniversalTime().ToString("HH:mm:ss.fff")}";
            est_1.Cells[13, edc.Type_Offset] = "I";
            est_1.Cells[13, edc.Name_Offset] = "Est8";
            est_1.Cells[13, edc.Distribution_Offset] = "Normal";
            est_1.Cells[13, edc.Distribution_Offset + 1] = 0;
            est_1.Cells[13, edc.Distribution_Offset + 2] = 1;

            System.Threading.Thread.Sleep(1);
            est_1.Cells[14, edc.ID_Offset] = $"DH|E|{ThisAddIn.MyApp.UserName}|{DateTime.Now.ToUniversalTime().ToString("ddMMyy")}{DateTime.Now.ToUniversalTime().ToString("HH:mm:ss.fff")}";
            est_1.Cells[14, edc.Type_Offset] = "I";
            est_1.Cells[14, edc.Name_Offset] = "Est9";
            est_1.Cells[14, edc.Distribution_Offset] = "Lognormal";
            est_1.Cells[14, edc.Distribution_Offset + 1] = 0;
            est_1.Cells[14, edc.Distribution_Offset + 2] = 1;

            System.Threading.Thread.Sleep(1);
            est_1.Cells[15, edc.ID_Offset] = $"DH|E|{ThisAddIn.MyApp.UserName}|{DateTime.Now.ToUniversalTime().ToString("ddMMyy")}{ DateTime.Now.ToUniversalTime().ToString("HH:mm:ss.fff")}";
            est_1.Cells[15, edc.Level_Offset] = 1;
            est_1.Cells[15, edc.Type_Offset] = "CE";
            est_1.Cells[15, edc.Name_Offset] = "Est10";
            est_1.Cells[15, edc.Distribution_Offset] = "Normal";
            est_1.Cells[15, edc.Distribution_Offset + 1] = 0;
            est_1.Cells[15, edc.Distribution_Offset + 2] = 1;


            for (int i = 0; i < 50; i++)
            {
                System.Threading.Thread.Sleep(1);
                est_1.Cells[16 + i, edc.ID_Offset] = $"DH|E|{ThisAddIn.MyApp.UserName}|{DateTime.Now.ToUniversalTime().ToString("ddMMyy")}{ DateTime.Now.ToUniversalTime().ToString("HH:mm:ss.fff")}";
                est_1.Cells[16 + i, edc.Type_Offset] = "I";
                est_1.Cells[16 + i, edc.Name_Offset] = $"Est{11+i}";
                est_1.Cells[16 + i, edc.Distribution_Offset] = "Normal";
                est_1.Cells[16 + i, edc.Distribution_Offset + 1] = 0;
                est_1.Cells[16 + i, edc.Distribution_Offset + 2] = 1;
            }

            est_1.Activate();

            wbs_1.Cells[4, wdc.ID_Offset] = "ID";
            wbs_1.Cells[4, wdc.Level_Offset] = "Level";
            wbs_1.Cells[4, wdc.Name_Offset] = "Name";
            wbs_1.Cells[4, wdc.Distribution_Offset] = "Distribution";
            wbs_1.Cells[4, wdc.Distribution_Offset + 1] = "Param1";
            wbs_1.Cells[4, wdc.Distribution_Offset + 2] = "Param2";
            wbs_1.Cells[4, wdc.Distribution_Offset + 3] = "Param3";

            System.Threading.Thread.Sleep(1);
            wbs_1.Cells[5, wdc.ID_Offset] = $"DH|S|{ThisAddIn.MyApp.UserName}|{DateTime.Now.ToUniversalTime().ToString("ddMMyy")}{DateTime.Now.ToUniversalTime().ToString("HH: mm:ss.fff")}";
            wbs_1.Cells[5, wdc.Level_Offset] = 1;
            wbs_1.Cells[5, wdc.Type_Offset] = "S";

            System.Threading.Thread.Sleep(1);
            wbs_1.Cells[6, wdc.ID_Offset] = $"DH|W|{ThisAddIn.MyApp.UserName}|{DateTime.Now.ToUniversalTime().ToString("ddMMyy")}{DateTime.Now.ToUniversalTime().ToString("HH: mm:ss.fff")}";
            wbs_1.Cells[6, wdc.Level_Offset] = 2;
            wbs_1.Cells[6, wdc.Type_Offset] = "CE";
            wbs_1.Cells[6, wdc.Name_Offset] = "Est2";
            wbs_1.Cells[6, wdc.Distribution_Offset] = "Triangular";
            wbs_1.Cells[6, wdc.Distribution_Offset + 1] = 10;
            wbs_1.Cells[6, wdc.Distribution_Offset + 2] = 30;
            wbs_1.Cells[6, wdc.Distribution_Offset + 3] = 20;

            System.Threading.Thread.Sleep(1);
            wbs_1.Cells[7, wdc.ID_Offset] = $"DH|W|{ThisAddIn.MyApp.UserName}|{DateTime.Now.ToUniversalTime().ToString("ddMMyy")}{DateTime.Now.ToUniversalTime().ToString("HH: mm:ss.fff")}";
            wbs_1.Cells[7, wdc.Level_Offset] = 2;
            wbs_1.Cells[7, wdc.Type_Offset] = "CE";
            wbs_1.Cells[7, wdc.Name_Offset] = "Est3";
            wbs_1.Cells[7, wdc.Distribution_Offset] = "Triangular";
            wbs_1.Cells[7, wdc.Distribution_Offset + 1] = 10;
            wbs_1.Cells[7, wdc.Distribution_Offset + 2] = 30;
            wbs_1.Cells[7, wdc.Distribution_Offset + 3] = 20;
 
            System.Threading.Thread.Sleep(1);
            wbs_1.Cells[8, wdc.ID_Offset] = $"DH|W|{ThisAddIn.MyApp.UserName}|{DateTime.Now.ToUniversalTime().ToString("ddMMyy")}{DateTime.Now.ToUniversalTime().ToString("HH: mm:ss.fff")}";
            wbs_1.Cells[8, wdc.Level_Offset] = 2;
            wbs_1.Cells[8, wdc.Type_Offset] = "CE";
            wbs_1.Cells[8, wdc.Name_Offset] = "Est4";
            wbs_1.Cells[8, wdc.Distribution_Offset] = "Triangular";
            wbs_1.Cells[8, wdc.Distribution_Offset + 1] = 10;
            wbs_1.Cells[8, wdc.Distribution_Offset + 2] = 30;
            wbs_1.Cells[8, wdc.Distribution_Offset + 3] = 20;

            System.Threading.Thread.Sleep(1);
            wbs_1.Cells[9, wdc.ID_Offset] = $"DH|W|{ThisAddIn.MyApp.UserName}|{DateTime.Now.ToUniversalTime().ToString("ddMMyy")}{DateTime.Now.ToUniversalTime().ToString("HH: mm:ss.fff")}";
            wbs_1.Cells[9, wdc.Level_Offset] = 2;
            wbs_1.Cells[9, wdc.Type_Offset] = "CE";
            wbs_1.Cells[9, wdc.Name_Offset] = "Est5";
            wbs_1.Cells[9, wdc.Distribution_Offset] = "Normal";
            wbs_1.Cells[9, wdc.Distribution_Offset + 1] = 0;
            wbs_1.Cells[9, wdc.Distribution_Offset + 2] = 1;

            System.Threading.Thread.Sleep(1);
            wbs_1.Cells[10, wdc.ID_Offset] = $"DH|S|{ThisAddIn.MyApp.UserName}|{DateTime.Now.ToUniversalTime().ToString("ddMMyy")}{DateTime.Now.ToUniversalTime().ToString("HH: mm:ss.fff")}";
            wbs_1.Cells[10, wdc.Level_Offset] = 1;
            wbs_1.Cells[10, wdc.Type_Offset] = "S";

            System.Threading.Thread.Sleep(1);
            wbs_1.Cells[11, wdc.ID_Offset] = $"DH|W|{ThisAddIn.MyApp.UserName}|{DateTime.Now.ToUniversalTime().ToString("ddMMyy")}{DateTime.Now.ToUniversalTime().ToString("HH: mm:ss.fff")}";
            wbs_1.Cells[11, wdc.Level_Offset] = 2;
            wbs_1.Cells[11, wdc.Type_Offset] = "CE";
            wbs_1.Cells[11, wdc.Name_Offset] = "Est6";
            wbs_1.Cells[11, wdc.Distribution_Offset] = "Normal";
            wbs_1.Cells[11, wdc.Distribution_Offset + 1] = 0;
            wbs_1.Cells[11, wdc.Distribution_Offset + 2] = 1;

            System.Threading.Thread.Sleep(1);
            wbs_1.Cells[12, wdc.ID_Offset] = $"DH|W|{ThisAddIn.MyApp.UserName}|{DateTime.Now.ToUniversalTime().ToString("ddMMyy")}{DateTime.Now.ToUniversalTime().ToString("HH: mm:ss.fff")}";
            wbs_1.Cells[12, wdc.Level_Offset] = 3;
            wbs_1.Cells[12, wdc.Type_Offset] = "CE";
            wbs_1.Cells[12, wdc.Name_Offset] = "Est7";
            wbs_1.Cells[12, wdc.Distribution_Offset] = "Normal";
            wbs_1.Cells[12, wdc.Distribution_Offset + 1] = 0;
            wbs_1.Cells[12, wdc.Distribution_Offset + 2] = 1;

            System.Threading.Thread.Sleep(1);
            wbs_1.Cells[13, wdc.ID_Offset] = $"DH|W|{ThisAddIn.MyApp.UserName}|{DateTime.Now.ToUniversalTime().ToString("ddMMyy")}{DateTime.Now.ToUniversalTime().ToString("HH: mm:ss.fff")}";
            wbs_1.Cells[13, wdc.Level_Offset] = 3;
            wbs_1.Cells[13, wdc.Type_Offset] = "CE";
            wbs_1.Cells[13, wdc.Name_Offset] = "Est8";
            wbs_1.Cells[13, wdc.Distribution_Offset] = "Normal";
            wbs_1.Cells[13, wdc.Distribution_Offset + 1] = 0;
            wbs_1.Cells[13, wdc.Distribution_Offset + 2] = 1;

            System.Threading.Thread.Sleep(1);
            wbs_1.Cells[14, wdc.ID_Offset] = $"DH|W|{ThisAddIn.MyApp.UserName}|{DateTime.Now.ToUniversalTime().ToString("ddMMyy")}{DateTime.Now.ToUniversalTime().ToString("HH: mm:ss.fff")}";
            wbs_1.Cells[14, wdc.Level_Offset] = 2;
            wbs_1.Cells[14, wdc.Type_Offset] = "CE";
            wbs_1.Cells[14, wdc.Name_Offset] = "Est9";
            wbs_1.Cells[14, wdc.Distribution_Offset] = "Lognormal";
            wbs_1.Cells[14, wdc.Distribution_Offset + 1] = 0;
            wbs_1.Cells[14, wdc.Distribution_Offset + 2] = 1;

            System.Threading.Thread.Sleep(1);
            wbs_1.Cells[15, wdc.ID_Offset] = $"DH|S|{ThisAddIn.MyApp.UserName}|{DateTime.Now.ToUniversalTime().ToString("ddMMyy")}{DateTime.Now.ToUniversalTime().ToString("HH: mm:ss.fff")}";
            wbs_1.Cells[15, wdc.Level_Offset] = 1;
            wbs_1.Cells[15, wdc.Type_Offset] = "S";

            System.Threading.Thread.Sleep(1);
            wbs_1.Cells[16, wdc.ID_Offset] = $"DH|W|{ThisAddIn.MyApp.UserName}|{DateTime.Now.ToUniversalTime().ToString("ddMMyy")}{DateTime.Now.ToUniversalTime().ToString("HH: mm:ss.fff")}";
            wbs_1.Cells[16, wdc.Level_Offset] = 2;
            wbs_1.Cells[16, wdc.Type_Offset] = "CE";
            wbs_1.Cells[16, wdc.Name_Offset] = "Est11";
            wbs_1.Cells[16, wdc.Distribution_Offset] = "Normal";
            wbs_1.Cells[16, wdc.Distribution_Offset + 1] = 0;
            wbs_1.Cells[16, wdc.Distribution_Offset + 2] = 1;

            //Goal: Build the correlation strings on each example sheet
            //Steps
            //1 -- Build the sheet object -- est_1 is the xlSheet; construct the sheet object from it
            
            CostSheet estimateSheet_example = CostSheet.ConstructFromXlCostSheet(est_1);

            //2 -- Manually load the estimate objects to the sheet object, including their SubEstimates
            estimateSheet_example.LoadCorrelStrings();
            estimateSheet_example.PrintDefaultCorrelStrings();
            
            //3 -- Build default CorrelStrings for estimates attached to the sheet object

            //Repeat for wbs_1
            CostSheet wbsSheet_example = CostSheet.ConstructFromXlCostSheet(wbs_1);
            wbsSheet_example.PrintDefaultCorrelStrings();

        }

        private void btnVisualize_Click(object sender, RibbonControlEventArgs e)
        {
            Sheets.CorrelationSheet correlSheet = Sheets.CorrelationSheet.ConstructFromXlCorrelationSheet();
            if (correlSheet == null)
                return;

            correlSheet.VisualizeCorrel();
        }

    }
}
