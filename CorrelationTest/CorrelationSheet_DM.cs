﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace CorrelationTest
{
    namespace Sheets
    {
        public class CorrelationSheet_DM : CorrelationSheet
        {
            public Data.CorrelationString_DM CorrelString { get; set; }

            public CorrelationSheet_DM(IHasDurationCorrelations ParentItem)        //bring in the coordinates and set up the ranges once they exist
            {   //Build from the correlString to get the xlSheet
                this.CorrelString = (Data.CorrelationString_DM)ParentItem.DurationCorrelationString;
                this.Specs = new Data.CorrelSheetSpecs(SheetType.Correlation_DM);
                this.xlSheet = GetXlSheet();
                this.LinkToOrigin = new Data.Link(ParentItem.xlCorrelCell_Duration);
                this.xlLinkCell = xlSheet.Cells[Specs.LinkCoords.Item1, Specs.LinkCoords.Item2];
                this.xlCorrelStringCell = xlSheet.Cells[Specs.StringCoords.Item1, Specs.StringCoords.Item2];
                this.xlIDCell = xlSheet.Cells[Specs.IdCoords.Item1, Specs.IdCoords.Item2];      //Is this junk?
                this.xlSubIdCell = xlSheet.Cells[Specs.SubIdCoords.Item1, Specs.SubIdCoords.Item2];
                this.xlDistCell = xlSheet.Cells[Specs.DistributionCoords.Item1, Specs.DistributionCoords.Item2];
                this.xlMatrixCell = xlSheet.Cells[Specs.MatrixCoords.Item1, Specs.MatrixCoords.Item2];
                this.xlPairsCell = xlSheet.Cells[Specs.PairsCoords.Item1, Specs.PairsCoords.Item2];
                this.Specs.PrintMatrixCoords(xlSheet);                                          //Print the matrix start coords
                this.Specs.PrintLinkCoords(xlSheet);                                            //Print the link coords
                this.Specs.PrintIdCoords(xlSheet);                                              //Print the ID coords
                this.Specs.PrintDistCoords(xlSheet);                                            //Print the Distribution coords
                CorrelMatrix = Data.CorrelationMatrix.ConstructFromParentItem(ParentItem, SheetType.Correlation_DM, this);
                this.PrintMatrixEndCoords(xlSheet);                                             //Print the matrix end coords
            }
            //COLLAPSE
            public CorrelationSheet_DM() //build from the xlsheet to get the string
            {
                //Need a link
                this.xlSheet = GetXlSheet();
                this.Specs = new Data.CorrelSheetSpecs(SheetType.Correlation_DM);
                //Set up the xlCells
                this.xlLinkCell = xlSheet.Cells[Specs.LinkCoords.Item1, Specs.LinkCoords.Item2];
                this.xlCorrelStringCell = xlSheet.Cells[Specs.StringCoords.Item1, Specs.StringCoords.Item2];
                this.xlIDCell = xlSheet.Cells[Specs.IdCoords.Item1, Specs.IdCoords.Item2];
                this.xlDistCell = xlSheet.Cells[Specs.DistributionCoords.Item1, Specs.DistributionCoords.Item2];
                this.xlSubIdCell = xlSheet.Cells[Specs.SubIdCoords.Item1, Specs.SubIdCoords.Item2];
                this.xlMatrixCell = xlSheet.Cells[Specs.MatrixCoords.Item1, Specs.MatrixCoords.Item2];
                this.xlPairsCell = xlSheet.Cells[Specs.PairsCoords.Item1, Specs.PairsCoords.Item2];

                //LINK
                this.LinkToOrigin = new Data.Link(xlLinkCell.Value);

                //Build the CorrelMatrix
                object[,] fieldsValues = xlSheet.Range[xlMatrixCell, xlMatrixCell.End[Excel.XlDirection.xlToRight]].Value;
                object[] ids = this.GetIDs();

                fieldsValues = ExtensionMethods.ReIndexArray(fieldsValues);
                object[] fields = ExtensionMethods.ToJaggedArray(fieldsValues)[0];
                Excel.Range matrixRange = xlSheet.Range[xlMatrixCell.Offset[1, 0], xlMatrixCell.End[Excel.XlDirection.xlToRight].End[Excel.XlDirection.xlDown]];
                object[,] matrix = matrixRange.Value;
                this.CorrelMatrix = Data.CorrelationMatrix.ConstructFromCorrelationSheet(this);
                //Build the CorrelString, which can print itself during collapse
                //Get these from the Header.
                //string parent_id = Convert.ToString(xlIDCell.Value);        //Get this from the header
                string parent_id = Data.CorrelationString.GetParentIDFromCorrelStringValue(xlCorrelStringCell.Value);
                SheetType sheetType = ExtensionMethods.GetSheetType(xlSheet);
                this.CorrelString = new Data.CorrelationString_DM(parent_id, ids, fields, this.CorrelMatrix);
            }

            //CONVERT
            public CorrelationSheet_DM(object[,] matrix, object[] ids, object[] fields, object header, object link, Excel.Worksheet replaceXlSheet) //build from the xlsheet to get the string
            {
                ThisAddIn.MyApp.DisplayAlerts = false;
                replaceXlSheet.Delete();
                ThisAddIn.MyApp.DisplayAlerts = true;
                this.xlSheet = this.GetXlSheet();
                this.Specs = new Data.CorrelSheetSpecs(SheetType.Correlation_DM);
                //Set up the xlCells
                this.xlLinkCell = this.xlSheet.Cells[this.Specs.LinkCoords.Item1, this.Specs.LinkCoords.Item2];
                this.xlCorrelStringCell = this.xlSheet.Cells[this.Specs.StringCoords.Item1, this.Specs.StringCoords.Item2];
                this.xlIDCell = this.xlSheet.Cells[this.Specs.IdCoords.Item1, this.Specs.IdCoords.Item2];
                this.xlDistCell = this.xlSheet.Cells[this.Specs.DistributionCoords.Item1, this.Specs.DistributionCoords.Item2];
                this.xlSubIdCell = this.xlSheet.Cells[this.Specs.SubIdCoords.Item1, this.Specs.SubIdCoords.Item2];
                this.xlMatrixCell = this.xlSheet.Cells[this.Specs.MatrixCoords.Item1, this.Specs.MatrixCoords.Item2];

                //LINK
                this.LinkToOrigin = new Data.Link(link.ToString());

                //Build the CorrelMatrix
                Excel.Range matrixRange = this.xlSheet.Range[this.xlMatrixCell.Offset[1, 0], this.xlMatrixCell.End[Excel.XlDirection.xlToRight].End[Excel.XlDirection.xlDown]];
                //This needs to construct off the un-printed sheet
                this.CorrelMatrix = Data.CorrelationMatrix.ConstructForConversion(matrix, ids, fields, header);
                //Build the CorrelString, which can print itself during collapse
                //Get these from the Header.
                //string parent_id = Convert.ToString(xlIDCell.Value);        //Get this from the header
                string parent_id = Data.CorrelationString.GetParentIDFromCorrelStringValue(header);
                SheetType sheetType = ExtensionMethods.GetSheetType(this.xlSheet);
                this.CorrelString = new Data.CorrelationString_DM(parent_id, ids, fields, this.CorrelMatrix);
            }

            protected override Excel.Worksheet GetXlSheet(bool CreateNew = true)        //Is this method being used?
            {
                var xlCorrelSheets = from Excel.Worksheet sheet in ThisAddIn.MyApp.Worksheets
                                     where sheet.Cells[1, 1].Value == "$CORRELATION_DM"
                                     select sheet;
                if (xlCorrelSheets.Any())
                    xlSheet = xlCorrelSheets.First();
                else if (CreateNew)
                    xlSheet = CreateXLCorrelSheet("_DM");
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

            protected override string GetDistributionString(IHasCorrelations est, int subIndex)
            {
                StringBuilder sb = new StringBuilder();
                sb.Append($"{((IHasDurationCorrelations)est).SubEstimates[subIndex].DurationDistribution.Name}");
                for (int i = 1; i < ((IHasDurationCorrelations)est).SubEstimates[subIndex].ValueDistributionParameters.Count(); i++)
                {
                    string param = $"Param{i}";
                    if (((IHasDurationCorrelations)est).SubEstimates[subIndex].ValueDistributionParameters[param] != null)
                        sb.Append($",{((IHasDurationCorrelations)est).SubEstimates[subIndex].ValueDistributionParameters[param]}");
                }
                return sb.ToString();
            }

            protected string GetSubIdString(IHasSubs est, int subIndex)
            {
                return ((IHasDurationCorrelations)est).SubEstimates[subIndex].uID.ID;
            }

            public override void PrintToSheet()  //expanding from string
            {
                //build a sheet object off the linksource
                CostSheet costSheet = CostSheet.ConstructFromXlCostSheet(this.LinkToOrigin.LinkSource.Worksheet);
                //Estimate_Item tempEst = new Estimate_Item(this.LinkToOrigin.LinkSource.EntireRow, costSheet);        //Load only this parent estimate
                //Aren't the estimates already here?
                Item parentItem = (from Item i in costSheet.Items where i.uID.ID == Convert.ToString(this.LinkToOrigin.LinkSource.EntireRow.Cells[1, costSheet.Specs.ID_Offset].value) select i).First();
                IHasDurationCorrelations parentEstimate;
                if (!(parentItem is IHasDurationCorrelations))
                {
                    throw new Exception("Item has no duration correlations");
                }
                else
                {
                    parentEstimate = (IHasDurationCorrelations)parentItem;
                }
                //IHasDurationCorrelations tempEst = (IHasDurationCorrelations)Item.ConstructFromRow(this.LinkToOrigin.LinkSource.EntireRow, costSheet);
                //tempEst.SubEstimates = tempEst.ContainingSheetObject.GetSubEstimates(tempEst.xlRow);                //Load the sub-estimates for this estimate
                this.CorrelMatrix.PrintToSheet(xlMatrixCell);                                   //Print the matrix
                this.LinkToOrigin.PrintToSheet(xlLinkCell);                                     //Print the link
                //This should really just print the header:
                CorrelString.PrintToSheet(xlCorrelStringCell);
                for (int subIndex = 0; subIndex < parentEstimate.SubEstimates.Count(); subIndex++)      //Print the Distribution strings
                {
                    //Distributions
                    if (parentEstimate.SubEstimates[subIndex] is IHasDurationCorrelations)
                        this.xlDistCell.Offset[subIndex, 0].Value = GetDistributionString(parentEstimate, subIndex);
                    //IDs
                    this.xlSubIdCell.Offset[subIndex, 0].Value = GetSubIdString(parentEstimate, subIndex);
                    this.xlSubIdCell.Offset[subIndex, 0].NumberFormat = "\"ID\";;;\"ID\"";
                }
                PrintColumnHeaders();
            }

            private void PrintColumnHeaders()
            {
                SheetType sType = ExtensionMethods.GetSheetType(this.xlSheet);
                if (sType == SheetType.Correlation_DP)
                {
                    this.xlPairsCell.Offset[-1, 0].Value = "Off-diagonal Values";
                    this.xlPairsCell.Offset[-1, 1].Value = "Linear reduction";
                }
                this.xlSubIdCell.Offset[-1, 0].Value = "Unique ID";

            }

            public override void FormatSheet()
            {
                Excel.Range matrixStart = this.xlMatrixCell.Offset[1, 0];
                Excel.Range matrixRange = matrixStart.Resize[this.CorrelMatrix.Fields.Length, this.CorrelMatrix.Fields.Length];

                //THIS SHOULD HAVE A DIFFERENT SUBCLASS
                if (ExtensionMethods.GetSheetType(this.xlSheet) == SheetType.Correlation_DM)
                {
                    foreach (Excel.Range cell in matrixRange.Cells)
                    {
                        int rowIndex = cell.Row - matrixStart.Row;
                        int colIndex = cell.Column - matrixStart.Column;
                        if (colIndex > rowIndex)
                            cell.Interior.Color = System.Drawing.Color.FromArgb(255, 255, 190);
                        else
                            cell.Interior.Color = System.Drawing.Color.FromArgb(225, 225, 225);
                    }
                }
                else if (ExtensionMethods.GetSheetType(this.xlSheet) == SheetType.Correlation_DP)
                {
                    foreach (Excel.Range cell in matrixRange.Cells)
                    {
                        int rowIndex = cell.Row - matrixStart.Row;
                        int colIndex = cell.Column - matrixStart.Column;
                        cell.Interior.Color = System.Drawing.Color.FromArgb(225, 225, 225);
                    }
                    Excel.Range xlPairsRange = this.xlPairsCell.Resize[matrixRange.Rows.Count - 1, 2];
                    foreach (Excel.Range cell in xlPairsRange)
                    {
                        cell.Interior.Color = System.Drawing.Color.FromArgb(255, 255, 190);
                    }
                }
            }

            public override void ConvertCorrelation(bool PreserveOffDiagonal=false)
            {
                /*
                 * This method needs to construct a _DM type using the information on the _DP type.
                 * This includes fitting the pairs to a matrix.
                 * Need the fields, matrix, IDs, Link, Header
                 */
                var pairs = PairSpecification.ConstructByFittingMatrix(this.CorrelMatrix.GetMatrix());
                object[] fields = this.GetFieldsFromXlCorrelSheet();
                object header = this.xlCorrelStringCell.Value;
                object link = this.xlLinkCell.Value;
                CorrelationSheet_DP convertedSheet = new CorrelationSheet_DP(pairs, fields, header, link);
                convertedSheet.PrintToSheet();
            }

            public override bool Validate() //This needs moved to subclass because the CorrelString implementation was moved to subclass
            {
                bool validateMatrix_to_String = this.CorrelString.ValidateAgainstMatrix(this.CorrelMatrix.Fields);
                //need to get fields from xlSheet fresh, not the object, to validate
                bool validateMatrix_to_xlSheet = this.CorrelMatrix.ValidateAgainstXlSheet(this.Get_xlFields());
                return validateMatrix_to_String && validateMatrix_to_xlSheet;
            }
        }
    }
}
