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
            public CorrelationSheet_Phasing(IHasPhasingCorrelations ParentItem)        //bring in the coordinates and set up the ranges once they exist
            {   //Build from the correlString to get the xlSheet
                this.CorrelString = ParentItem.PhasingCorrelationString;
                SheetType correlType = CorrelString.GetCorrelType();
                this.Specs = new Data.CorrelSheetSpecs(correlType);
                this.xlSheet = GetXlSheet(correlType);
                CorrelMatrix = Data.CorrelationMatrix.ConstructFromParentItem(ParentItem, correlType);
                this.LinkToOrigin = new Data.Link(ParentItem.xlCorrelCell_Phasing);
                this.xlLinkCell = xlSheet.Cells[Specs.LinkCoords.Item1, Specs.LinkCoords.Item2];
                this.xlCorrelStringCell = xlSheet.Cells[Specs.StringCoords.Item1, Specs.StringCoords.Item2];
                this.xlIDCell = xlSheet.Cells[Specs.IdCoords.Item1, Specs.IdCoords.Item2];
                this.xlSubIdCell = xlSheet.Cells[Specs.SubIdCoords.Item1, Specs.SubIdCoords.Item2];
                this.xlDistCell = xlSheet.Cells[Specs.DistributionCoords.Item1, Specs.DistributionCoords.Item2];
                this.xlMatrixCell = xlSheet.Cells[Specs.MatrixCoords.Item1, Specs.MatrixCoords.Item2];
                this.xlPairsCell = xlSheet.Cells[Specs.TripleCoords.Item1, Specs.TripleCoords.Item2];
                this.Specs.PrintMatrixCoords(xlSheet);                                          //Print the matrix start coords
                this.PrintMatrixEndCoords(xlSheet);                                             //Print the matrix end coords
                this.Specs.PrintLinkCoords(xlSheet);                                            //Print the link coords
                this.Specs.PrintIdCoords(xlSheet);                                              //Print the ID coords
                this.Specs.PrintDistCoords(xlSheet);                                            //Print the Distribution coords
            }

            public CorrelationSheet_Phasing(SheetType shtType) //build from the xlsheet to get the string
            {
                this.xlSheet = GetXlSheet(shtType);
                this.Specs = new Data.CorrelSheetSpecs(shtType);
                //Set up the xlCells
                this.xlLinkCell = xlSheet.Cells[Specs.LinkCoords.Item1, Specs.LinkCoords.Item2];
                this.xlCorrelStringCell = xlSheet.Cells[Specs.StringCoords.Item1, Specs.StringCoords.Item2];
                this.xlIDCell = xlSheet.Cells[Specs.IdCoords.Item1, Specs.IdCoords.Item2];
                this.xlDistCell = xlSheet.Cells[Specs.DistributionCoords.Item1, Specs.IdCoords.Item2];
                this.xlMatrixCell = xlSheet.Cells[Specs.MatrixCoords.Item1, Specs.MatrixCoords.Item2];
                this.xlPairsCell = xlSheet.Cells[Specs.TripleCoords.Item1, Specs.TripleCoords.Item2];

                //Set up the link
                this.LinkToOrigin = new Data.Link(Convert.ToString(xlLinkCell.Value));

                //Build the CorrelMatrix
                int fields = Convert.ToInt32(Convert.ToString(xlCorrelStringCell.Value).Split(',')[0]);
                Excel.Range fieldRange = xlMatrixCell.Resize[1, fields];
                Excel.Range matrixRange = xlMatrixCell.Offset[1, 0].Resize[fields, fields];
                //this.CorrelMatrix = new Data.CorrelationMatrix(this, fieldRange, matrixRange);
                this.CorrelMatrix = Data.CorrelationMatrix.ConstructFromCorrelationSheet(this);
                //Build the CorrelString, which can print itself during collapse
                SheetType sheetType = ExtensionMethods.GetSheetType(xlSheet);
                if (sheetType == SheetType.Correlation_PM)
                    this.CorrelString = new Data.CorrelationString_PM(this.CorrelMatrix, Convert.ToString(this.xlIDCell.Value));
                else if (sheetType == SheetType.Correlation_PP)
                {
                    //Build the triple from the string
                    string correlStringVal = this.xlCorrelStringCell.Value;
                    Data.CorrelationString_PP existing_cst = new Data.CorrelationString_PP(correlStringVal);
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

            //public override void UpdateCorrelationString(string[] ids)
            //{
            //    this.CorrelString = new Data.CorrelationString_PM(ids, this.CorrelMatrix);
            //}

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
                        case SheetType.Correlation_PP:
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

            public override void UpdateCorrelationString(string[] ids)
            {
                UniqueID parentID = UniqueID.ConstructFromExisting(Convert.ToString(this.xlIDCell.Value));
                object[,] matrix = this.xlMatrixCell.Offset[1, 0].Resize[ids.Length, ids.Length].Value;
                this.CorrelMatrix = Data.CorrelationMatrix.ConstructFromCorrelationSheet(this);
                this.CorrelString = new Data.CorrelationString_CM(parentID.ID, ids, this.CorrelMatrix.Fields, CorrelMatrix);
                this.xlCorrelStringCell.Value = this.CorrelString.Value;
            }

            //Bring in the coordinates - use an enum to build them for each sheet type
            //Parse CorrelString to get type for collapse

            protected override string GetDistributionString(IHasCorrelations est)
            {
                StringBuilder sb = new StringBuilder();
                sb.Append($"{((IHasPhasingCorrelations)est).PhasingDistribution.Name}");
                for (int i = 1; i < ((IHasPhasingCorrelations)est).PhasingDistribution.DistributionParameters.Count(); i++)
                {
                    string param = $"Param{i}";
                    if (((IHasPhasingCorrelations)est).PhasingDistribution.DistributionParameters[param] != null)
                        sb.Append($",{((IHasPhasingCorrelations)est).PhasingDistribution.DistributionParameters[param]}");
                }
                return sb.ToString();
            }

            public override void PrintToSheet()  //expanding from string
            {
                //build a sheet object off the linksource
                CostSheet costSheet = CostSheet.ConstructFromXlCostSheet(this.LinkToOrigin.LinkSource.Worksheet);
                IHasPhasingCorrelations tempEst = (IHasPhasingCorrelations)Item.ConstructFromRow(this.LinkToOrigin.LinkSource.EntireRow, costSheet);        //Load only this parent estimate
                //tempEst.LoadSubEstimates();                //Load the sub-estimates for this estimate
                //tempEst.ContainingSheetObject.GetSubEstimates(tempEst.xlRow);                //Load the sub-estimates for this estimate
                this.CorrelMatrix.PrintToSheet(xlMatrixCell);                                   //Print the matrix
                this.LinkToOrigin.PrintToSheet(xlLinkCell);                                     //Print the link
                this.xlIDCell.Value = tempEst.uID.ID;                                               //Print the ID
                this.xlIDCell.ColumnWidth = 40;
                CorrelString.PrintToSheet(xlCorrelStringCell);
                this.xlDistCell.Value = GetDistributionString(tempEst);

                if (CorrelString is Data.CorrelationString_PP)       //Need to replicate this in PT and DT.
                {
                    this.xlPairsCell.Value = ((Data.CorrelationString_PP)CorrelString).GetTriple().GetValuesString();
                }
                //this.xlDistCell.Value = tempEst.ItemDistribution.Name;
            }
        }
    }
}
