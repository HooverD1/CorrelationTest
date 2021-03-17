﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using Vsto = Microsoft.Office.Tools.Excel;

namespace CorrelationTest
{
    namespace Sheets
    {
        public class CorrelationSheet_CM : CorrelationSheet
        {
            public Data.CorrelationString_CM CorrelString { get; set; }
            public Excel.Range xlButton_ConvertCorrel { get; set; }

            //EXPAND
            public CorrelationSheet_CM(IHasCostCorrelations ParentItem)        //bring in the coordinates and set up the ranges once they exist
            {
                this.CorrelString = (Data.CorrelationString_CM)ParentItem.CostCorrelationString;
                SheetType correlType = CorrelString.GetCorrelType();
                this.Specs = new Data.CorrelSheetSpecs(correlType);
                this.xlSheet = GetXlSheet();

                this.LinkToOrigin = new Data.Link(ParentItem.xlCorrelCell_Cost);
                this.xlLinkCell = xlSheet.Cells[Specs.LinkCoords.Item1, Specs.LinkCoords.Item2];
                this.xlCorrelStringCell = xlSheet.Cells[Specs.StringCoords.Item1, Specs.StringCoords.Item2];
                //this.xlIDCell = xlSheet.Cells[Specs.IdCoords.Item1, Specs.IdCoords.Item2];
                this.xlSubIdCell = xlSheet.Cells[Specs.SubIdCoords.Item1, Specs.SubIdCoords.Item2];
                this.xlDistCell = xlSheet.Cells[Specs.DistributionCoords.Item1, Specs.DistributionCoords.Item2];
                this.xlMatrixCell = xlSheet.Cells[Specs.MatrixCoords.Item1, Specs.MatrixCoords.Item2];
                this.xlPairsCell = xlSheet.Cells[Specs.PairsCoords.Item1, Specs.PairsCoords.Item2];
                this.Specs.PrintMatrixCoords(xlSheet);                                          //Print the matrix start coords

                this.Specs.PrintLinkCoords(xlSheet);                                            //Print the link coords
                this.Specs.PrintIdCoords(xlSheet);                                              //Print the ID coords
                this.Specs.PrintDistCoords(xlSheet);                                            //Print the Distribution coords
                CorrelMatrix = Data.CorrelationMatrix.ConstructFromParentItem(ParentItem, correlType, this);
                this.PrintMatrixEndCoords(xlSheet);                                             //Print the matrix end coords
                this.FormatSheet();
            }

            //COLLAPSE METHOD
            public CorrelationSheet_CM() //build from the xlsheet to get the string
            {
                this.xlSheet = GetXlSheet();
                this.Specs = new Data.CorrelSheetSpecs(SheetType.Correlation_CM);
                //Set up the xlCells
                this.xlLinkCell = xlSheet.Cells[Specs.LinkCoords.Item1, Specs.LinkCoords.Item2];
                this.xlCorrelStringCell = xlSheet.Cells[Specs.StringCoords.Item1, Specs.StringCoords.Item2];
                //this.xlIDCell = xlSheet.Cells[Specs.IdCoords.Item1, Specs.IdCoords.Item2];
                this.xlDistCell = xlSheet.Cells[Specs.DistributionCoords.Item1, Specs.DistributionCoords.Item2];
                this.xlSubIdCell = xlSheet.Cells[Specs.SubIdCoords.Item1, Specs.SubIdCoords.Item2];
                this.xlMatrixCell = xlSheet.Cells[Specs.MatrixCoords.Item1, Specs.MatrixCoords.Item2];
                this.LinkToOrigin = new Data.Link(Convert.ToString(xlLinkCell.Value));
                //
                //Build the CorrelMatrix
                object[] ids = Data.CorrelationString.GetIDsFromString(xlCorrelStringCell.Value);
                object[,] fieldsValues = xlSheet.Range[xlMatrixCell, xlMatrixCell.End[Excel.XlDirection.xlToRight]].Value;
                fieldsValues = ExtensionMethods.ReIndexArray(fieldsValues);
                object[] fields = ExtensionMethods.ToJaggedArray(fieldsValues)[0];
                Excel.Range matrixRange = xlSheet.Range[xlMatrixCell.Offset[1, 0], xlMatrixCell.End[Excel.XlDirection.xlDown].End[Excel.XlDirection.xlToRight]];
                object[,] matrix = matrixRange.Value;
                this.CorrelMatrix = Data.CorrelationMatrix.ConstructFromCorrelationSheet(this);
                //Build the CorrelString, which can print itself during collapse
                string parent_id = Data.CorrelationString.GetParentIDFromCorrelStringValue(xlCorrelStringCell.Value);
                SheetType sheetType = ExtensionMethods.GetSheetType(xlSheet);
                if (sheetType == SheetType.Correlation_CP)
                {
                    //Build the triple from the string
                    //Need to build the CorrelationString_CP without access to its string.
                    //Header -- follow link & build parent/subs, use their IDs
                    //Old Triple values -- print to sheet when expanding from a triple
                    //Values -- Build the Matrix

                    //string correlStringVal = this.xlCorrelStringCell.Value;
                    //Data.CorrelationString_CP existing_cst = new Data.CorrelationString_CP(correlStringVal);

                    PairSpecification pairs = PairSpecification.ConstructFromString(Convert.ToString(xlPairsCell.Value));
                    //Check if the matrix still matches the triple.
                    if (this.CorrelMatrix.ValidateAgainstPairs(pairs))
                    {       //If YES - create cs_triple object
                        this.CorrelString = (Data.CorrelationString_CM)Data.CorrelationString.ConstructFromCorrelationSheet(this);
                    }
                    else
                    {       //If NO - create cs_periods object
                        this.CorrelString = new Data.CorrelationString_CM(parent_id, ids, fields, CorrelMatrix);
                    }
                }
                else if (sheetType == SheetType.Correlation_CM)
                {
                    this.CorrelString = new Data.CorrelationString_CM(parent_id, ids, fields, this.CorrelMatrix);
                }
                this.FormatSheet();
            }

            //CONVERT
            public CorrelationSheet_CM(object[,] matrix, object[] ids, object[] fields, object header, object link, Excel.Worksheet replaceXlSheet) //build from the xlsheet to get the string
            {
                ThisAddIn.MyApp.DisplayAlerts = false;
                replaceXlSheet.Delete();
                ThisAddIn.MyApp.DisplayAlerts = true;
                this.xlSheet = this.GetXlSheet();
                this.Specs = new Data.CorrelSheetSpecs(SheetType.Correlation_CM);
                //Set up the xlCells
                this.xlLinkCell = this.xlSheet.Cells[this.Specs.LinkCoords.Item1, this.Specs.LinkCoords.Item2];
                this.xlCorrelStringCell = this.xlSheet.Cells[this.Specs.StringCoords.Item1, this.Specs.StringCoords.Item2];
                //this.xlIDCell = this.xlSheet.Cells[this.Specs.IdCoords.Item1, this.Specs.IdCoords.Item2];
                this.xlDistCell = this.xlSheet.Cells[this.Specs.DistributionCoords.Item1, this.Specs.DistributionCoords.Item2];
                this.xlSubIdCell = this.xlSheet.Cells[this.Specs.SubIdCoords.Item1, this.Specs.SubIdCoords.Item2];
                this.xlMatrixCell = this.xlSheet.Cells[this.Specs.MatrixCoords.Item1, this.Specs.MatrixCoords.Item2];
                this.xlButton_ConvertCorrel = xlSheet.Cells[Specs.Btn_ConvertCoords.Item1, Specs.Btn_ConvertCoords.Item2];
                //LINK
                this.LinkToOrigin = new Data.Link(link.ToString());

                //Build the CorrelMatrix
                Excel.Range matrixRange = this.xlSheet.Range[this.xlMatrixCell.Offset[1, 0], this.xlMatrixCell.End[Excel.XlDirection.xlToRight].End[Excel.XlDirection.xlDown]];
                //This needs to construct off the un-printed sheet
                this.CorrelMatrix = Data.CorrelationMatrix.ConstructForConversion(matrix, ids, fields, header);
                //Build the CorrelString, which can print itself during collapse
                string parent_id = Data.CorrelationString.GetParentIDFromCorrelStringValue(header);
                SheetType sheetType = ExtensionMethods.GetSheetType(this.xlSheet);
                this.CorrelString = new Data.CorrelationString_CM(parent_id, ids, fields, this.CorrelMatrix);
                this.Header = header.ToString();
            }


            public override void FormatSheet()
            {
                Excel.Range matrixStart = this.xlMatrixCell.Offset[1, 0];
                Excel.Range matrixRange = matrixStart.Resize[this.CorrelMatrix.Fields.Length, this.CorrelMatrix.Fields.Length];
                Excel.Range diagonal = matrixRange.Cells[1, 1];
                for (int i = 2; i <= matrixRange.Columns.Count; i++)
                {
                    diagonal = ThisAddIn.MyApp.Union(diagonal, matrixRange.Cells[i, i]);
                }

                foreach (Excel.Range cell in matrixRange.Cells)
                {
                    int rowIndex = cell.Row - matrixStart.Row;
                    int colIndex = cell.Column - matrixStart.Column;
                    if (colIndex > rowIndex)
                        cell.Interior.Color = System.Drawing.Color.FromArgb(255, 255, 190);
                    else
                        cell.Interior.Color = System.Drawing.Color.FromArgb(225, 225, 225);
                }

                diagonal.Interior.Color = System.Drawing.Color.FromArgb(0, 0, 0);
                diagonal.Font.Color = System.Drawing.Color.FromArgb(255, 255, 255);

                matrixRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                matrixRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            }

            protected override Excel.Worksheet GetXlSheet(bool CreateNew = true)
            {
                var xlCorrelSheets = from Excel.Worksheet sheet in ThisAddIn.MyApp.Worksheets
                                     where sheet.Cells[1, 1].Value == "$CORRELATION_CM"
                                     select sheet;
                if (xlCorrelSheets.Any())
                    xlSheet = xlCorrelSheets.First();
                else if (CreateNew)
                    xlSheet = CreateXLCorrelSheet("_CM");
                else
                    throw new Exception("No input matrix correlation sheet found.");
                return xlSheet;
            }

            protected override Excel.Worksheet CreateXLCorrelSheet(string postfix)
            {
                Excel.Worksheet xlCorrelSheet = ThisAddIn.MyApp.Worksheets.Add(After: ThisAddIn.MyApp.ActiveWorkbook.Sheets[ThisAddIn.MyApp.ActiveWorkbook.Sheets.Count]);
                xlCorrelSheet.Name = "Correlation";
                xlCorrelSheet.Cells[1, 1] = $"$CORRELATION{postfix}";
                xlCorrelSheet.Rows[1].Hidden = true;
                return xlCorrelSheet;
            }

            public override void UpdateCorrelationString(string[] ids)
            {
                //What is the purpose of this method? To update the matrix values?
                //Update the string on the sheet to match the altered matrix
                UniqueID parentID = UniqueID.ConstructFromExisting(Data.CorrelationString.GetParentIDFromCorrelStringValue(xlCorrelStringCell.Value));
                object[,] matrix = this.xlMatrixCell.Offset[1, 0].Resize[ids.Length, ids.Length].Value;
                this.CorrelMatrix = Data.CorrelationMatrix.ConstructFromCorrelationSheet(this);
                this.CorrelString = new Data.CorrelationString_CM(parentID.ID, ids, this.CorrelMatrix.Fields, CorrelMatrix);
                this.xlCorrelStringCell.Value = this.CorrelString.Value;
            }

            protected override string GetDistributionString(IHasCorrelations est, int subIndex)
            {
                StringBuilder sb = new StringBuilder();
                sb.Append($"{((IHasCostCorrelations)est).SubEstimates[subIndex].CostDistribution.Name}");
                for (int i = 1; i < ((IHasCostCorrelations)est).SubEstimates[subIndex].ValueDistributionParameters.Count(); i++)
                {
                    string param = $"Param{i}";
                    if (((IHasCostCorrelations)est).SubEstimates[subIndex].ValueDistributionParameters[param] != null)
                        sb.Append($",{((IHasCostCorrelations)est).SubEstimates[subIndex].ValueDistributionParameters[param]}");
                }
                return sb.ToString();
            }

            protected string GetSubIdString(IHasSubs est, int subIndex)
            {
                return ((IHasCostCorrelations)est).SubEstimates[subIndex].uID.ID;
            }

            private void AddUserControls()
            {
                Vsto.Worksheet vstoSheet = Globals.Factory.GetVstoObject(this.xlSheet);
                System.Windows.Forms.Button btn_ConvertToCP = new System.Windows.Forms.Button();
                btn_ConvertToCP.Text = "Convert to Matrix Specification";
                btn_ConvertToCP.Click += ConversionFormClicked;
                vstoSheet.Controls.AddControl(btn_ConvertToCP, this.xlButton_ConvertCorrel.Resize[1, 3], "ConvertToCP");
            }



            public override void PrintToSheet()  //expanding from string
            {
                //build a sheet object off the linksource
                CostSheet costSheet = CostSheet.ConstructFromXlCostSheet(this.LinkToOrigin.LinkSource.Worksheet);
                //Estimate_Item tempEst = new Estimate_Item(this.LinkToOrigin.LinkSource.EntireRow, costSheet);        //Load only this parent estimate

                //This needs to find the parent..
                //IHasCostCorrelations tempEst = (IHasCostCorrelations)Item.ConstructFromRow(this.LinkToOrigin.LinkSource.EntireRow, costSheet);
                //tempEst.SubEstimates = tempEst.ContainingSheetObject.GetSubEstimates(tempEst.xlRow);                //Load the sub-estimates for this estimate

                Item parentItem = (from Item i in costSheet.Items where i.uID.ID == Convert.ToString(this.LinkToOrigin.LinkSource.EntireRow.Cells[1, costSheet.Specs.ID_Offset].value) select i).First();
                IHasCostCorrelations parentEstimate;

                if (!(parentItem is IHasCostCorrelations))
                {
                    throw new Exception("Item has no cost correlations");
                }
                else
                {
                    parentEstimate = (IHasCostCorrelations)parentItem;
                }

                this.CorrelMatrix.PrintToSheet(xlMatrixCell);                                   //Print the matrix
                this.LinkToOrigin.PrintToSheet(xlLinkCell);                                     //Print the link

                CorrelString.PrintToSheet(xlCorrelStringCell);

                for (int subIndex = 0; subIndex < parentEstimate.SubEstimates.Count(); subIndex++)      //Print the Distribution strings
                {
                    //Distributions
                    this.xlDistCell.Offset[subIndex, 0].Value = GetDistributionString(parentEstimate, subIndex);
                    //IDs
                    this.xlSubIdCell.Offset[subIndex, 0].Value = GetSubIdString(parentEstimate, subIndex);
                    this.xlSubIdCell.Offset[subIndex, 0].NumberFormat = "\"ID\";;;\"ID\"";
                }
                PrintColumnHeaders();
                AddUserControls();
            }

            private void PrintColumnHeaders()
            {
                SheetType sType = ExtensionMethods.GetSheetType(this.xlSheet);
                if (sType == SheetType.Correlation_CP)
                {
                    this.xlPairsCell.Offset[-1, 0].Value = "Off-diagonal Values";
                    this.xlPairsCell.Offset[-1, 1].Value = "Linear reduction";
                }
                this.xlSubIdCell.Offset[-1, 0].Value = "Unique ID";

            }

            public override bool Validate() //This needs moved to subclass because the CorrelString implementation was moved to subclass
            {
                bool validateMatrix_to_String = this.CorrelString.ValidateAgainstMatrix(this.CorrelMatrix.Fields);
                //need to get fields from xlSheet fresh, not the object, to validate
                bool validateMatrix_to_xlSheet = this.CorrelMatrix.ValidateAgainstXlSheet(this.Get_xlFields());
                return validateMatrix_to_String && validateMatrix_to_xlSheet;
            }

            public override void ConvertCorrelation(bool PreserveOffDiagonal=false)
            {
                /*
                 * This method needs to construct a _DM type using the information on the _DP type.
                 * This includes fitting the pairs to a matrix.
                 * Need the fields, matrix, IDs, Link, Header
                 */
                var pairs = PairSpecification.ConstructByFittingMatrix(this.CorrelMatrix.GetMatrix_Values(), PreserveOffDiagonal);
                object[] ids = this.GetIDs();
                object[] fields = this.GetFields();
                object header = this.xlCorrelStringCell.Value;
                object link = this.xlLinkCell.Value;
                CorrelationSheet_CP convertedSheet = new CorrelationSheet_CP(pairs, ids, fields, header, link, this.xlSheet);
                convertedSheet.PrintToSheet();
                convertedSheet.FormatSheet();
            }
        }
    }
}
