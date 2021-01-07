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
        public class CorrelationSheet_Inputs : CorrelationSheet
        {            
            public CorrelationSheet_Inputs(Data.CorrelationString_IM correlString, Excel.Range launchedFrom, Data.CorrelSheetSpecs specs)        //bring in the coordinates and set up the ranges once they exist
            {   //Build from the correlString to get the xlSheet
                this.CorrelString = correlString;
                this.Specs = specs;
                this.xlSheet = GetXlSheet();
                CorrelMatrix = new Data.CorrelationMatrix((Data.CorrelationString_IM)CorrelString);
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

            public CorrelationSheet_Inputs(Data.CorrelationString_IT correlString, Excel.Range launchedFrom, Data.CorrelSheetSpecs specs)        //bring in the coordinates and set up the ranges once they exist
            {   //Build from the correlString to get the xlSheet
                this.CorrelString = correlString;
                this.Specs = specs;
                this.xlSheet = GetXlSheet();
                CorrelMatrix = new Data.CorrelationMatrix((Data.CorrelationString_IT)CorrelString);
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

            public CorrelationSheet_Inputs(Data.CorrelSheetSpecs specs) //build from the xlsheet to get the string
            {
                this.xlSheet = GetXlSheet(false);
                this.Specs = specs;
                //Set up the xlCells
                this.xlLinkCell = xlSheet.Cells[specs.LinkCoords.Item1, specs.LinkCoords.Item2];
                this.xlCorrelStringCell = xlSheet.Cells[specs.StringCoords.Item1, specs.StringCoords.Item2];
                this.xlIDCell = xlSheet.Cells[specs.IdCoords.Item1, specs.IdCoords.Item2];
                this.xlDistCell = xlSheet.Cells[specs.DistributionCoords.Item1, specs.IdCoords.Item2];
                this.xlMatrixCell = xlSheet.Cells[specs.MatrixCoords.Item1, specs.MatrixCoords.Item2];
                //
                //Build the CorrelMatrix
                Excel.Range matrixRange = xlSheet.Range[xlMatrixCell.Offset[1,0], xlMatrixCell.End[Excel.XlDirection.xlDown].End[Excel.XlDirection.xlToRight]];
                this.CorrelMatrix = new Data.CorrelationMatrix(this, xlMatrixCell.Resize[1,matrixRange.Columns.Count], matrixRange);  
                //Build the CorrelString, which can print itself during collapse
                this.CorrelString = new Data.CorrelationString_IM(this.CorrelMatrix);
            }

            protected override Excel.Worksheet GetXlSheet(bool CreateNew = true)
            {
                var xlCorrelSheets = from Excel.Worksheet sheet in ThisAddIn.MyApp.Worksheets
                                     where sheet.Cells[1, 1].Value == "$CORRELATION_IM"
                                     select sheet;
                if (xlCorrelSheets.Any())
                    xlSheet = xlCorrelSheets.First();
                else if (CreateNew)
                    xlSheet = CreateXLCorrelSheet("_IM");
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

            public override void UpdateCorrelationString()
            {
                this.CorrelString = new Data.CorrelationString_IM(this.CorrelMatrix);
            }

            public override void PrintToSheet()  //expanding from string
            {
                //build a sheet object off the linksource
                CostSheet costSheet = CostSheet.Construct(this.LinkToOrigin.LinkSource.Worksheet);
                Estimate_Item tempEst = new Estimate_Item(this.LinkToOrigin.LinkSource.EntireRow, costSheet);        //Load only this parent estimate
                tempEst.ContainingSheetObject.GetSubEstimates(tempEst.xlRow);                //Load the sub-estimates for this estimate
                this.CorrelMatrix.PrintToSheet(xlMatrixCell);                                   //Print the matrix
                this.LinkToOrigin.PrintToSheet(xlLinkCell);                                     //Print the link
                this.xlIDCell.Value = tempEst.uID.ID;                                               //Print the ID
                CorrelString.PrintToSheet(xlCorrelStringCell);
                for (int subIndex = 0; subIndex < tempEst.SubEstimates.Count(); subIndex++)      //Print the Distribution strings
                {
                    this.xlDistCell.Offset[subIndex, 0].Value = GetDistributionString(tempEst, subIndex);
                }
            }
        }
    }
}
