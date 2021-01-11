using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace CorrelationTest
{
    namespace Sheets
    {
        public class CorrelationSheet_Phasing : CorrelationSheet
        {
            public CorrelationSheet_Phasing(Data.CorrelationString_PM correlString, Excel.Range launchedFrom, Data.CorrelSheetSpecs specs)        //bring in the coordinates and set up the ranges once they exist
            {
                this.CorrelString = correlString;
                this.Specs = specs;
                var xlCorrelSheets = from Excel.Worksheet sheet in ThisAddIn.MyApp.Worksheets
                                     where sheet.Cells[1, 1].Value == "$CORRELATION_PM"
                                     select sheet;
                if (xlCorrelSheets.Any())
                    xlSheet = xlCorrelSheets.First();
                else
                    xlSheet = CreateXLCorrelSheet("_PM");
                CorrelMatrix = Data.CorrelationMatrix.ConstructNew((Data.CorrelationString_PM)CorrelString);
                this.LinkToOrigin = new Data.Link(launchedFrom);
                this.xlLinkCell = xlSheet.Cells[specs.LinkCoords.Item1, specs.LinkCoords.Item2];
                this.xlCorrelStringCell = xlSheet.Cells[specs.StringCoords.Item1, specs.StringCoords.Item2];
                this.xlIDCell = xlSheet.Cells[specs.IdCoords.Item1, specs.IdCoords.Item2];
                this.xlDistCell = xlSheet.Cells[specs.DistributionCoords.Item1, specs.IdCoords.Item2];
                this.xlMatrixCell = xlSheet.Cells[specs.MatrixCoords.Item1, specs.MatrixCoords.Item2];
                this.Specs.PrintMatrixCoords(xlSheet);                                          //Print the matrix start coords
                this.PrintMatrixEndCoords(xlSheet);                                             //Print the matrix end coords
                this.Specs.PrintLinkCoords(xlSheet);                                            //Print the link coords
                this.Specs.PrintIdCoords(xlSheet);                                              //Print the ID coords
                this.Specs.PrintDistCoords(xlSheet);                                            //Print the Distribution coords
            }

            public CorrelationSheet_Phasing(Data.CorrelationString_PT correlString, Excel.Range launchedFrom, Data.CorrelSheetSpecs specs)        //bring in the coordinates and set up the ranges once they exist
            {
                this.CorrelString = correlString;
                this.Specs = specs;
                var xlCorrelSheets = from Excel.Worksheet sheet in ThisAddIn.MyApp.Worksheets
                                     where sheet.Cells[1, 1].Value == "$CORRELATION_PT"
                                     select sheet;
                if (xlCorrelSheets.Any())
                    xlSheet = xlCorrelSheets.First();
                else
                    xlSheet = CreateXLCorrelSheet("_PT");
                CorrelMatrix = Data.CorrelationMatrix.ConstructNew((Data.CorrelationString_PT)CorrelString);
                this.LinkToOrigin = new Data.Link(launchedFrom);
                this.xlLinkCell = xlSheet.Cells[specs.LinkCoords.Item1, specs.LinkCoords.Item2];
                this.xlCorrelStringCell = xlSheet.Cells[specs.StringCoords.Item1, specs.StringCoords.Item2];
                this.xlIDCell = xlSheet.Cells[specs.IdCoords.Item1, specs.IdCoords.Item2];
                this.xlDistCell = xlSheet.Cells[specs.DistributionCoords.Item1, specs.IdCoords.Item2];
                this.xlMatrixCell = xlSheet.Cells[specs.MatrixCoords.Item1, specs.MatrixCoords.Item2];
                this.Specs.PrintMatrixCoords(xlSheet);                                          //Print the matrix start coords
                this.PrintMatrixEndCoords(xlSheet);                                             //Print the matrix end coords
                this.Specs.PrintLinkCoords(xlSheet);                                            //Print the link coords
                this.Specs.PrintIdCoords(xlSheet);                                              //Print the ID coords
                this.Specs.PrintDistCoords(xlSheet);                                            //Print the Distribution coords
            }

            public CorrelationSheet_Phasing(Data.CorrelSheetSpecs specs) //build from the xlsheet to get the string
            {
                this.xlSheet = GetXlSheet();
                this.Specs = specs;
                //Set up the xlCells
                this.xlLinkCell = xlSheet.Cells[specs.LinkCoords.Item1, specs.LinkCoords.Item2];
                this.xlCorrelStringCell = xlSheet.Cells[specs.StringCoords.Item1, specs.StringCoords.Item2];
                this.xlIDCell = xlSheet.Cells[specs.IdCoords.Item1, specs.IdCoords.Item2];
                this.xlDistCell = xlSheet.Cells[specs.DistributionCoords.Item1, specs.IdCoords.Item2];
                this.xlMatrixCell = xlSheet.Cells[specs.MatrixCoords.Item1, specs.MatrixCoords.Item2];
                //
                //Build the CorrelMatrix
                int fields = Convert.ToInt32(Convert.ToString(xlCorrelStringCell.Value).Split(',')[0]);
                Excel.Range fieldRange = xlMatrixCell.Resize[1, fields];
                Excel.Range matrixRange = xlMatrixCell.Offset[1, 0].Resize[fields, fields];
                this.CorrelMatrix = new Data.CorrelationMatrix(this, fieldRange, matrixRange);
                //Build the CorrelString, which can print itself during collapse
                SheetType sheetType = ExtensionMethods.GetSheetType(xlSheet);
                if (sheetType == SheetType.Correlation_PM)
                    this.CorrelString = new Data.CorrelationString_PM(this.CorrelMatrix, Convert.ToString(this.xlIDCell.Value));
                else if (sheetType == SheetType.Correlation_PT)
                {
                    //Build the triple from the string
                    string correlStringVal = this.xlCorrelStringCell.Value;
                    Data.CorrelationString_PT existing_cst = new Data.CorrelationString_PT(correlStringVal);
                    Triple pt = existing_cst.GetTriple();
                    //Check if the matrix still matches the triple.
                    if (this.CorrelMatrix.ValidateAgainstTriple(pt))
                    {       //If YES - create cs_triple object
                        this.CorrelString = existing_cst;
                    }
                    else
                    {       //If NO - create cs_periods object
                        this.CorrelString = new Data.CorrelationString_PM(this.CorrelMatrix, Convert.ToString(this.xlIDCell.Value));
                    }
                }
                else
                {
                    throw new Exception("Invalid sheet type.");
                }
            }

            protected Excel.Worksheet GetXlSheet() { return GetXlSheet(false); }

            protected override Excel.Worksheet GetXlSheet(bool CreateNew = true)
            {
                var xlCorrelSheets = from Excel.Worksheet sheet in ThisAddIn.MyApp.Worksheets
                                     where sheet.Cells[1, 1].Value == "$CORRELATION_PM" || sheet.Cells[1, 1].value == "$CORRELATION_PT"
                                     select sheet;
                if (xlCorrelSheets.Any())
                    xlSheet = xlCorrelSheets.First();
                else
                    throw new Exception("No correlation sheet found.");
                return xlSheet;
            }

            protected override Excel.Worksheet GetXlSheet(SheetType sheetType, bool CreateNew = true)
            {
                var xlCorrelSheets = from Excel.Worksheet sheet in ThisAddIn.MyApp.Worksheets
                                     where sheet.Cells[1, 1].Value == "$CORRELATION_PM" || sheet.Cells[1,1].value == "$CORRELATION_PT"
                                     select sheet;
                if (xlCorrelSheets.Any())
                    xlSheet = xlCorrelSheets.First();
                else if (CreateNew)
                {
                    switch (sheetType)
                    {
                        case SheetType.Correlation_PM:
                            xlSheet = CreateXLCorrelSheet("_PM");
                            break;
                        case SheetType.Correlation_PT:
                            xlSheet = CreateXLCorrelSheet("_PT");
                            break;
                        default:
                            throw new Exception("Bad sheet type");
                    }                    
                }                    
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

            //Bring in the coordinates - use an enum to build them for each sheet type
            //Parse CorrelString to get type for collapse

            public override void PrintToSheet()  //expanding from string
            {
                //build a sheet object off the linksource
                CostSheet costSheet = CostSheet.Construct(this.LinkToOrigin.LinkSource.Worksheet);
                Estimate_Item tempEst = new Estimate_Item(this.LinkToOrigin.LinkSource.EntireRow, costSheet);        //Load only this parent estimate
                //tempEst.LoadSubEstimates();                //Load the sub-estimates for this estimate
                this.CorrelMatrix.PrintToSheet(xlMatrixCell);                                   //Print the matrix
                this.LinkToOrigin.PrintToSheet(xlLinkCell);                                     //Print the link
                this.xlIDCell.Value = tempEst.uID.ID;                                               //Print the ID
                CorrelString.PrintToSheet(xlCorrelStringCell);
                this.xlDistCell.Value = tempEst.ItemDistribution.Name;
            }
        }
    }
}
