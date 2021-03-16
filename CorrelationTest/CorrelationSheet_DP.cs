using System;
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
        public class CorrelationSheet_DP : CorrelationSheet
        {
            public PairSpecification PairSpec { get; set; }
            public Data.CorrelationString_DP CorrelString { get; set; }
            public Excel.Range xlButton_ConvertToDM { get; set; }

            public CorrelationSheet_DP(IHasDurationCorrelations ParentItem)        //bring in the coordinates and set up the ranges once they exist
            {   //Build from the correlString to get the xlSheet
                this.CorrelString = (Data.CorrelationString_DP)ParentItem.DurationCorrelationString;
                this.Specs = new Data.CorrelSheetSpecs(SheetType.Correlation_DP);
                this.xlSheet = GetXlSheet();
                this.LinkToOrigin = new Data.Link(ParentItem.xlCorrelCell_Duration);
                this.xlLinkCell = xlSheet.Cells[Specs.LinkCoords.Item1, Specs.LinkCoords.Item2];
                this.xlCorrelStringCell = xlSheet.Cells[Specs.StringCoords.Item1, Specs.StringCoords.Item2];
                this.xlIDCell = xlSheet.Cells[Specs.IdCoords.Item1, Specs.IdCoords.Item2];      //Is this junk?
                this.xlSubIdCell = xlSheet.Cells[Specs.SubIdCoords.Item1, Specs.SubIdCoords.Item2];
                this.xlDistCell = xlSheet.Cells[Specs.DistributionCoords.Item1, Specs.DistributionCoords.Item2];
                this.xlMatrixCell = xlSheet.Cells[Specs.MatrixCoords.Item1, Specs.MatrixCoords.Item2];
                this.xlPairsCell = xlSheet.Cells[Specs.PairsCoords.Item1, Specs.PairsCoords.Item2];
                this.xlButton_ConvertToDM = xlSheet.Cells[Specs.Btn_ConvertCoords.Item1, Specs.Btn_ConvertCoords.Item2];
                this.Specs.PrintMatrixCoords(xlSheet);                                          //Print the matrix start coords
                this.Specs.PrintLinkCoords(xlSheet);                                            //Print the link coords
                this.Specs.PrintIdCoords(xlSheet);                                              //Print the ID coords
                this.Specs.PrintDistCoords(xlSheet);                                            //Print the Distribution coords
                CorrelMatrix = Data.CorrelationMatrix.ConstructFromParentItem(ParentItem, SheetType.Correlation_DP, this);
                this.PrintMatrixEndCoords(xlSheet);                                             //Print the matrix end coords

            }
            //COLLAPSE
            public CorrelationSheet_DP() //build from the xlsheet to get the string
            {
                //Need a link
                this.xlSheet = GetXlSheet();
                this.Specs = new Data.CorrelSheetSpecs(SheetType.Correlation_DP);
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
                object[] fields = this.GetFields();
                object[] ids = this.GetIDs();

                Excel.Range matrixRange = xlSheet.Range[xlMatrixCell.Offset[1, 0], xlMatrixCell.End[Excel.XlDirection.xlToRight].End[Excel.XlDirection.xlDown]];
                object[,] matrix = matrixRange.Value;
                this.CorrelMatrix = Data.CorrelationMatrix.ConstructFromCorrelationSheet(this);
                //Build the CorrelString, which can print itself during collapse
                //Get these from the Header.
                //string parent_id = Convert.ToString(xlIDCell.Value);        //Get this from the header
                string parent_id = Data.CorrelationString.GetParentIDFromCorrelStringValue(xlCorrelStringCell.Value);
                SheetType sheetType = ExtensionMethods.GetSheetType(xlSheet);
                PairSpecification pairs = PairSpecification.ConstructFromRange(xlPairsCell, fields.Count() - 1);

                //Check if the matrix still matches the triple.
                if (this.CorrelMatrix.ValidateAgainstPairs(pairs))
                {       //If YES - create cs_triple object
                    this.CorrelString = (Data.CorrelationString_DP)Data.CorrelationString.ConstructFromCorrelationSheet(this);
                }
                else
                {       //If NO - create cs_periods object
                    throw new NotImplementedException();
                    //Alert the user and give option to cancel or launch conversion form
                }                
            }

            //CONVERT
            public CorrelationSheet_DP(PairSpecification pairs, object[] fields, object header, object link)
            {
                //Need a link
                this.xlSheet = GetXlSheet();
                this.Specs = new Data.CorrelSheetSpecs(SheetType.Correlation_DP);
                //Set up the xlCells
                this.xlLinkCell = xlSheet.Cells[Specs.LinkCoords.Item1, Specs.LinkCoords.Item2];
                this.xlCorrelStringCell = xlSheet.Cells[Specs.StringCoords.Item1, Specs.StringCoords.Item2];
                this.xlIDCell = xlSheet.Cells[Specs.IdCoords.Item1, Specs.IdCoords.Item2];
                this.xlDistCell = xlSheet.Cells[Specs.DistributionCoords.Item1, Specs.DistributionCoords.Item2];
                this.xlSubIdCell = xlSheet.Cells[Specs.SubIdCoords.Item1, Specs.SubIdCoords.Item2];
                this.xlMatrixCell = xlSheet.Cells[Specs.MatrixCoords.Item1, Specs.MatrixCoords.Item2];
                this.xlPairsCell = xlSheet.Cells[Specs.PairsCoords.Item1, Specs.PairsCoords.Item2];

                //LINK
                this.LinkToOrigin = new Data.Link(link.ToString());

                //Build the CorrelMatrix
                object[,] fieldsValues = xlSheet.Range[xlMatrixCell, xlMatrixCell.End[Excel.XlDirection.xlToRight]].Value;
                object[] ids = ExtensionMethods.ToJaggedArray((object[,])this.xlSubIdCell.Resize[fieldsValues.GetLength(1), 1].Value)[1];

                fieldsValues = ExtensionMethods.ReIndexArray(fieldsValues);
                Excel.Range matrixRange = xlSheet.Range[xlMatrixCell.Offset[1, 0], xlMatrixCell.End[Excel.XlDirection.xlToRight].End[Excel.XlDirection.xlDown]];
                object[,] matrix = matrixRange.Value;
                this.CorrelMatrix = Data.CorrelationMatrix.ConstructFromCorrelationSheet(this);
                //Build the CorrelString, which can print itself during collapse
                //Get these from the Header.
                //string parent_id = Convert.ToString(xlIDCell.Value);        //Get this from the header
                string parent_id = Data.CorrelationString.GetParentIDFromCorrelStringValue(xlCorrelStringCell.Value);
                SheetType sheetType = ExtensionMethods.GetSheetType(xlSheet);
                this.PairSpec = pairs;
            }

            protected override Excel.Worksheet GetXlSheet(bool CreateNew = true)        //Is this method being used?
            {
                var xlCorrelSheets = from Excel.Worksheet sheet in ThisAddIn.MyApp.Worksheets
                                     where sheet.Cells[1, 1].value == "$CORRELATION_DP"
                                     select sheet;
                if (xlCorrelSheets.Any())
                    xlSheet = xlCorrelSheets.First();
                else if (CreateNew)
                    xlSheet = CreateXLCorrelSheet("_DP");
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

            //public override void UpdateCorrelationString(string[] ids)
            //{
            //    //What is the purpose of this method? To update the matrix values?
            //    //Update the string on the sheet to match the altered matrix
            //    UniqueID parentID = UniqueID.ConstructFromExisting(Convert.ToString(this.xlIDCell.Value));
            //    object[,] matrix = this.xlMatrixCell.Offset[1, 0].Resize[ids.Length, ids.Length].Value;
            //    this.CorrelMatrix = Data.CorrelationMatrix.ConstructFromCorrelationSheet(this);
            //    this.CorrelString = new Data.CorrelationString_CM(parentID.ID, ids, this.CorrelMatrix.Fields, CorrelMatrix);
            //    this.xlCorrelStringCell.Value = this.CorrelString.Value;
            //}

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
                this.xlPairsCell.Resize[parentEstimate.SubEstimates.Count() - 1, 2].Value = ((Data.CorrelationString_DP)CorrelString).GetPairwise().GetValuesString_Split();
                PrintColumnHeaders();
                AddUserControls();
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

            private void AddUserControls()
            {
                Vsto.Worksheet vstoSheet = Globals.Factory.GetVstoObject(this.xlSheet);
                System.Windows.Forms.Button btn_ConvertToDM = new System.Windows.Forms.Button();
                vstoSheet.Controls.AddControl(btn_ConvertToDM, this.xlButton_ConvertToDM.Resize[1,3], "ConvertToDM");
                btn_ConvertToDM.Text = "Convert to Matrix Specification";
                btn_ConvertToDM.Click += ConversionFormClicked;
            }

            private void ConversionFormClicked(object sender, EventArgs e)      //This works.. but why? Isn't the object gone?
            {
                var conversionForm = new CorrelationConversionForm();
                conversionForm.Show();
                conversionForm.Focus();
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
                object[,] matrix = this.CorrelMatrix.GetMatrix();
                object[] ids = this.GetIDs();   //This isn't returning anything
                object[] fields = this.CorrelMatrix.Fields;
                object header = this.xlCorrelStringCell.Value;
                object link = this.xlLinkCell.Value;
                Sheets.CorrelationSheet_DM newSheet = new Sheets.CorrelationSheet_DM(matrix, ids, fields, header, link, this.xlSheet);
                newSheet.PrintToSheet();
                newSheet.FormatSheet();
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
