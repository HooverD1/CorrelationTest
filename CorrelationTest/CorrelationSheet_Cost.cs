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
        public class CorrelationSheet_Cost : CorrelationSheet
        {
            public CorrelationSheet_Cost(IHasCostCorrelations ParentItem)        //bring in the coordinates and set up the ranges once they exist
            {   
                this.CorrelString = ParentItem.CostCorrelationString;
                SheetType correlType = CorrelString.GetCorrelType();
                this.Specs = new Data.CorrelSheetSpecs(correlType);
                this.xlSheet = GetXlSheet(correlType);
                
                this.LinkToOrigin = new Data.Link(ParentItem.xlCorrelCell_Cost);
                this.xlLinkCell = xlSheet.Cells[Specs.LinkCoords.Item1, Specs.LinkCoords.Item2];
                this.xlCorrelStringCell = xlSheet.Cells[Specs.StringCoords.Item1, Specs.StringCoords.Item2];
                this.xlIDCell = xlSheet.Cells[Specs.IdCoords.Item1, Specs.IdCoords.Item2];
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
            }

            //COLLAPSE METHOD
            public CorrelationSheet_Cost(SheetType shtType) //build from the xlsheet to get the string
            {
                this.xlSheet = GetXlSheet(shtType);
                this.Specs = new Data.CorrelSheetSpecs(shtType);
                //Set up the xlCells
                this.xlLinkCell = xlSheet.Cells[Specs.LinkCoords.Item1, Specs.LinkCoords.Item2];
                this.xlCorrelStringCell = xlSheet.Cells[Specs.StringCoords.Item1, Specs.StringCoords.Item2];
                this.xlIDCell = xlSheet.Cells[Specs.IdCoords.Item1, Specs.IdCoords.Item2];
                this.xlDistCell = xlSheet.Cells[Specs.DistributionCoords.Item1, Specs.DistributionCoords.Item2];
                this.xlSubIdCell = xlSheet.Cells[Specs.SubIdCoords.Item1, Specs.SubIdCoords.Item2];
                this.xlMatrixCell = xlSheet.Cells[Specs.MatrixCoords.Item1, Specs.MatrixCoords.Item2];
                if(shtType == SheetType.Correlation_CP)
                    this.xlPairsCell = xlSheet.Cells[Specs.PairsCoords.Item1, Specs.PairsCoords.Item2];
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
                //Get these from the Header.
                string parent_id = Convert.ToString(xlIDCell.Value);
                SheetType sheetType = ExtensionMethods.GetSheetType(xlSheet);
                if(sheetType == SheetType.Correlation_CP)
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
                        this.CorrelString = Data.CorrelationString.ConstructFromCorrelationSheet(this);
                    }
                    else
                    {       //If NO - create cs_periods object
                        this.CorrelString = new Data.CorrelationString_CM(parent_id, ids, fields, CorrelMatrix);
                    }
                }
                else if(sheetType == SheetType.Correlation_CM)
                {
                    this.CorrelString = new Data.CorrelationString_CM(parent_id, ids, fields, this.CorrelMatrix);
                }
                
            }
            
            protected override Excel.Worksheet GetXlSheet(SheetType sheetType, bool CreateNew = true)
            {
                var xlCorrelSheets = from Excel.Worksheet sheet in ThisAddIn.MyApp.Worksheets
                                     where sheet.Cells[1, 1].Value == "$CORRELATION_CM" || sheet.Cells[1,1].value == "$CORRELATION_CT"
                                     select sheet;
                if (xlCorrelSheets.Any())
                    xlSheet = xlCorrelSheets.First();
                else if (CreateNew)
                {
                    switch (sheetType)
                    {
                        case SheetType.Correlation_CM:
                            xlSheet = CreateXLCorrelSheet("_CM");
                            break;
                        case SheetType.Correlation_CP:
                            xlSheet = CreateXLCorrelSheet("_CP");
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
                //What is the purpose of this method? To update the matrix values?
                //Update the string on the sheet to match the altered matrix
                UniqueID parentID = UniqueID.ConstructFromExisting(Convert.ToString(this.xlIDCell.Value));
                object[,] matrix = this.xlMatrixCell.Offset[1,0].Resize[ids.Length, ids.Length].Value;
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

            public override void PrintToSheet()  //expanding from string
            {
                //build a sheet object off the linksource
                CostSheet costSheet = CostSheet.ConstructFromXlCostSheet(this.LinkToOrigin.LinkSource.Worksheet);
                //Estimate_Item tempEst = new Estimate_Item(this.LinkToOrigin.LinkSource.EntireRow, costSheet);        //Load only this parent estimate
                
                //This needs to find the parent..
                IHasCostCorrelations tempEst = (IHasCostCorrelations)Item.ConstructFromRow(this.LinkToOrigin.LinkSource.EntireRow, costSheet);
                tempEst.SubEstimates = tempEst.ContainingSheetObject.GetSubEstimates(tempEst.xlRow);                //Load the sub-estimates for this estimate
                this.CorrelMatrix.PrintToSheet(xlMatrixCell);                                   //Print the matrix
                this.LinkToOrigin.PrintToSheet(xlLinkCell);                                     //Print the link
                this.xlIDCell.Value = tempEst.uID.ID;                                               //Print the ID
                this.xlIDCell.ColumnWidth = 40;
                
                CorrelString.PrintToSheet(xlCorrelStringCell);
                if(CorrelString is Data.CorrelationString_CP)       //Need to replicate this in PT and DT.
                {
                    this.xlPairsCell.Value = ((Data.CorrelationString_CP)CorrelString).GetPairs().Value;
                }
                    
                for (int subIndex = 0; subIndex < tempEst.SubEstimates.Count(); subIndex++)      //Print the Distribution strings
                {
                    //Distributions
                    this.xlDistCell.Offset[subIndex, 0].Value = GetDistributionString(tempEst, subIndex);
                    //IDs
                    this.xlSubIdCell.Offset[subIndex, 0].Value = GetSubIdString(tempEst, subIndex);
                    this.xlSubIdCell.Offset[subIndex, 0].NumberFormat = "\"ID\";;;\"ID\"";
                }
            }
        }
    }
}
