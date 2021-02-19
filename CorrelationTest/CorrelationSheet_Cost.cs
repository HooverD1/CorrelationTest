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
            public CorrelationSheet_Cost(Data.CorrelationString_CM correlString, Excel.Range launchedFrom, Data.CorrelSheetSpecs specs)        //bring in the coordinates and set up the ranges once they exist
            {   //Build from the correlString to get the xlSheet
                this.CorrelString = correlString;
                this.Specs = specs;
                this.xlSheet = GetXlSheet();
                CorrelMatrix = Data.CorrelationMatrix.ConstructNew((Data.CorrelationString_CM)CorrelString);
                this.LinkToOrigin = new Data.Link(launchedFrom);
                this.xlLinkCell = xlSheet.Cells[specs.LinkCoords.Item1, specs.LinkCoords.Item2];
                this.xlCorrelStringCell = xlSheet.Cells[specs.StringCoords.Item1, specs.StringCoords.Item2];
                this.xlIDCell = xlSheet.Cells[specs.IdCoords.Item1, specs.IdCoords.Item2];
                this.xlSubIdCell = xlSheet.Cells[specs.SubIdCoords.Item1, specs.SubIdCoords.Item2];
                this.xlDistCell = xlSheet.Cells[specs.DistributionCoords.Item1, specs.DistributionCoords.Item2];
                this.xlMatrixCell = xlSheet.Cells[specs.MatrixCoords.Item1, specs.MatrixCoords.Item2];
                this.xlTripleCell = xlSheet.Cells[specs.TripleCoords.Item1, specs.TripleCoords.Item2];
                this.Specs.PrintMatrixCoords(xlSheet);                                          //Print the matrix start coords
                this.PrintMatrixEndCoords(xlSheet);                                             //Print the matrix end coords
                this.Specs.PrintLinkCoords(xlSheet);                                            //Print the link coords
                this.Specs.PrintIdCoords(xlSheet);                                              //Print the ID coords
                this.Specs.PrintDistCoords(xlSheet);                                            //Print the Distribution coords

            }

            public CorrelationSheet_Cost(Data.CorrelationString_CT correlString, Excel.Range launchedFrom, Data.CorrelSheetSpecs specs)        //bring in the coordinates and set up the ranges once they exist
            {   //Build from the correlString to get the xlSheet
                this.CorrelString = correlString;
                this.Specs = specs;
                this.xlSheet = GetXlSheet(SheetType.Correlation_CT);
                //This needs constructed from the parent item. The parent item needs to contain the string.
                CorrelMatrix = Data.CorrelationMatrix.ConstructNew((Data.CorrelationString_CT)CorrelString, GetFields());
                this.LinkToOrigin = new Data.Link(launchedFrom);
                this.xlLinkCell = xlSheet.Cells[specs.LinkCoords.Item1, specs.LinkCoords.Item2];
                this.xlCorrelStringCell = xlSheet.Cells[specs.StringCoords.Item1, specs.StringCoords.Item2];
                this.xlIDCell = xlSheet.Cells[specs.IdCoords.Item1, specs.IdCoords.Item2];
                this.xlSubIdCell = xlSheet.Cells[specs.SubIdCoords.Item1, specs.SubIdCoords.Item2];
                this.xlDistCell = xlSheet.Cells[specs.DistributionCoords.Item1, specs.DistributionCoords.Item2];
                this.xlMatrixCell = xlSheet.Cells[specs.MatrixCoords.Item1, specs.MatrixCoords.Item2];
                this.xlTripleCell = xlSheet.Cells[specs.TripleCoords.Item1, specs.TripleCoords.Item2];
                this.Specs.PrintMatrixCoords(xlSheet);                                          //Print the matrix start coords
                this.PrintMatrixEndCoords(xlSheet);                                             //Print the matrix end coords
                this.Specs.PrintLinkCoords(xlSheet);                                            //Print the link coords
                this.Specs.PrintIdCoords(xlSheet);                                              //Print the ID coords
                this.Specs.PrintDistCoords(xlSheet);                                            //Print the Distribution coords
            }

            public CorrelationSheet_Cost(SheetType shtType) //build from the xlsheet to get the string
            {
                this.xlSheet = GetXlSheet();
                this.Specs = new Data.CorrelSheetSpecs(shtType);
                //Set up the xlCells
                this.xlLinkCell = xlSheet.Cells[Specs.LinkCoords.Item1, Specs.LinkCoords.Item2];
                this.xlCorrelStringCell = xlSheet.Cells[Specs.StringCoords.Item1, Specs.StringCoords.Item2];
                this.xlIDCell = xlSheet.Cells[Specs.IdCoords.Item1, Specs.IdCoords.Item2];
                this.xlDistCell = xlSheet.Cells[Specs.DistributionCoords.Item1, Specs.DistributionCoords.Item2];
                this.xlSubIdCell = xlSheet.Cells[Specs.SubIdCoords.Item1, Specs.SubIdCoords.Item2];
                this.xlMatrixCell = xlSheet.Cells[Specs.MatrixCoords.Item1, Specs.MatrixCoords.Item2];
                this.xlTripleCell = xlSheet.Cells[Specs.TripleCoords.Item1, Specs.TripleCoords.Item2];
                this.LinkToOrigin = new Data.Link(xlLinkCell.Value);
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
                if(sheetType == SheetType.Correlation_CT)
                {
                    //Build the triple from the string
                    //Need to build the CorrelationString_CT without access to its string.
                    //Header -- follow link & build parent/subs, use their IDs
                    //Old Triple values -- print to sheet when expanding from a triple
                    //Values -- Build the Matrix

                    //string correlStringVal = this.xlCorrelStringCell.Value;
                    //Data.CorrelationString_CT existing_cst = new Data.CorrelationString_CT(correlStringVal);

                    Triple triple = new Triple(Convert.ToString(xlTripleCell.Value));
                    //Check if the matrix still matches the triple.
                    if (this.CorrelMatrix.ValidateAgainstTriple(triple))
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
            
            protected override Excel.Worksheet GetXlSheet(bool CreateNew = true)
            {
                var xlCorrelSheets = from Excel.Worksheet sheet in ThisAddIn.MyApp.Worksheets
                                     where sheet.Cells[1, 1].Value == "$CORRELATION_CM" || sheet.Cells[1,1].value == "$CORRELATION_CT"
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

            protected override Excel.Worksheet GetXlSheet(SheetType sheetType, bool CreateNew = true)
            {
                var xlCorrelSheets = from Excel.Worksheet sheet in ThisAddIn.MyApp.Worksheets
                                     where sheet.Cells[1, 1].Value == "$CORRELATION_CM" || sheet.Cells[1, 1].value == "$CORRELATION_CT"
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
                        case SheetType.Correlation_CT:
                            xlSheet = CreateXLCorrelSheet("_CT");
                            break;
                        default:
                            throw new Exception("Bad sheet type");
                    }
                }
                else
                    throw new Exception("No input matrix correlation sheet found.");
                return xlSheet;
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

            protected override string GetDistributionString(IHasSubs est, int subIndex)
            {
                StringBuilder sb = new StringBuilder();
                sb.Append($"{((IHasCostSubs)est).SubEstimates[subIndex].CostDistribution.Name}");
                for (int i = 1; i < ((IHasCostSubs)est).SubEstimates[subIndex].ValueDistributionParameters.Count(); i++)
                {
                    string param = $"Param{i}";
                    if (((IHasCostSubs)est).SubEstimates[subIndex].ValueDistributionParameters[param] != null)
                        sb.Append($",{((IHasCostSubs)est).SubEstimates[subIndex].ValueDistributionParameters[param]}");
                }
                return sb.ToString();
            }

            protected string GetSubIdString(IHasSubs est, int subIndex)
            {
                return ((IHasCostSubs)est).SubEstimates[subIndex].uID.ID;
            }

            public override void PrintToSheet()  //expanding from string
            {
                //build a sheet object off the linksource
                CostSheet costSheet = CostSheet.Construct(this.LinkToOrigin.LinkSource.Worksheet);
                //Estimate_Item tempEst = new Estimate_Item(this.LinkToOrigin.LinkSource.EntireRow, costSheet);        //Load only this parent estimate
                
                //This needs to find the parent..
                IHasCostSubs tempEst = (IHasCostSubs)Item.ConstructFromRow(this.LinkToOrigin.LinkSource.EntireRow, costSheet);
                tempEst.SubEstimates = tempEst.ContainingSheetObject.GetSubEstimates(tempEst.xlRow);                //Load the sub-estimates for this estimate
                this.CorrelMatrix.PrintToSheet(xlMatrixCell);                                   //Print the matrix
                this.LinkToOrigin.PrintToSheet(xlLinkCell);                                     //Print the link
                this.xlIDCell.Value = tempEst.uID.ID;                                               //Print the ID
                this.xlIDCell.ColumnWidth = 40;
                
                CorrelString.PrintToSheet(xlCorrelStringCell);
                if(CorrelString is Data.CorrelationString_CT)       //Need to replicate this in PT and DT.
                {
                    this.xlTripleCell.Value = ((Data.CorrelationString_CT)CorrelString).GetTriple().GetValuesString();
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
