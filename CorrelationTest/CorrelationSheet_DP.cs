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
            public Excel.Range xlButton_ConvertCorrel { get; set; }

            //EXPAND
            public CorrelationSheet_DP(IHasDurationCorrelations ParentItem)        //bring in the coordinates and set up the ranges once they exist
            {   //Build from the correlString to get the xlSheet
                this.CorrelString = (Data.CorrelationString_DP)ParentItem.DurationCorrelationString;
                this.Specs = new Data.CorrelSheetSpecs(SheetType.Correlation_DP);
                this.xlSheet = GetXlSheet();
                this.LinkToOrigin = new Data.Link(ParentItem.xlCorrelCell_Duration);
                this.xlLinkCell = xlSheet.Cells[Specs.LinkCoords.Item1, Specs.LinkCoords.Item2];
                this.xlHeaderCell = xlSheet.Cells[Specs.StringCoords.Item1, Specs.StringCoords.Item2];
                //this.xlIDCell = xlSheet.Cells[Specs.IdCoords.Item1, Specs.IdCoords.Item2];      //Is this junk?
                this.xlSubIdCell = xlSheet.Cells[Specs.SubIdCoords.Item1, Specs.SubIdCoords.Item2];
                this.xlDistCell = xlSheet.Cells[Specs.DistributionCoords.Item1, Specs.DistributionCoords.Item2];
                this.xlMatrixCell = xlSheet.Cells[Specs.MatrixCoords.Item1, Specs.MatrixCoords.Item2];
                this.xlPairsCell = xlSheet.Cells[Specs.PairsCoords.Item1, Specs.PairsCoords.Item2];
                this.xlButton_ConvertCorrel = xlSheet.Cells[Specs.Btn_ConvertCoords.Item1, Specs.Btn_ConvertCoords.Item2];
                this.xlButton_CollapseCorrel = xlSheet.Cells[Specs.Btn_Collapse.Item1, Specs.Btn_Collapse.Item2];
                this.xlButton_Visualize = xlSheet.Cells[Specs.Btn_Visualize.Item1, Specs.Btn_Visualize.Item2];
                this.xlButton_Cancel = xlSheet.Cells[Specs.Btn_Cancel.Item1, Specs.Btn_Cancel.Item2];
                CorrelMatrix = Data.CorrelationMatrix.ConstructFromParentItem(ParentItem, SheetType.Correlation_DP, this);
                this.Header = CorrelString.GetHeader();
                this.PairSpec = ((Data.CorrelationString_DP)CorrelString).GetPairwise();
            }
            //COLLAPSE
            public CorrelationSheet_DP() //build from the xlsheet to get the string
            {
                //Need a link
                this.xlSheet = GetXlSheet();
                this.Specs = new Data.CorrelSheetSpecs(SheetType.Correlation_DP);
                //Set up the xlCells
                this.xlLinkCell = xlSheet.Cells[Specs.LinkCoords.Item1, Specs.LinkCoords.Item2];
                this.xlHeaderCell = xlSheet.Cells[Specs.StringCoords.Item1, Specs.StringCoords.Item2];
                //this.xlIDCell = xlSheet.Cells[Specs.IdCoords.Item1, Specs.IdCoords.Item2];
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
                string parent_id = Data.CorrelationString.GetParentIDFromCorrelStringValue(xlHeaderCell.Value);
                SheetType sheetType = ExtensionMethods.GetSheetType(xlSheet);
                PairSpecification pairs = PairSpecification.ConstructFromRange(xlPairsCell, fields.Count() - 1);
            }

            //CONVERT
            public CorrelationSheet_DP(PairSpecification pairs, object[] ids, object[] fields, object header, object link, Excel.Worksheet replaceXlSheet)
            {
                ThisAddIn.MyApp.DisplayAlerts = false;
                replaceXlSheet.Delete();
                ThisAddIn.MyApp.DisplayAlerts = true;
                this.xlSheet = GetXlSheet();
                this.Specs = new Data.CorrelSheetSpecs(SheetType.Correlation_DP);
                //Set up the xlCells
                this.xlLinkCell = xlSheet.Cells[Specs.LinkCoords.Item1, Specs.LinkCoords.Item2];
                this.xlHeaderCell = xlSheet.Cells[Specs.StringCoords.Item1, Specs.StringCoords.Item2];
                //this.xlIDCell = xlSheet.Cells[Specs.IdCoords.Item1, Specs.IdCoords.Item2];
                this.xlDistCell = xlSheet.Cells[Specs.DistributionCoords.Item1, Specs.DistributionCoords.Item2];
                this.xlSubIdCell = xlSheet.Cells[Specs.SubIdCoords.Item1, Specs.SubIdCoords.Item2];
                this.xlMatrixCell = xlSheet.Cells[Specs.MatrixCoords.Item1, Specs.MatrixCoords.Item2];
                this.xlPairsCell = xlSheet.Cells[Specs.PairsCoords.Item1, Specs.PairsCoords.Item2];
                this.xlButton_ConvertCorrel = xlSheet.Cells[Specs.Btn_ConvertCoords.Item1, Specs.Btn_ConvertCoords.Item2];
                this.xlButton_CollapseCorrel = xlSheet.Cells[Specs.Btn_Collapse.Item1, Specs.Btn_Collapse.Item2];
                this.xlButton_Visualize = xlSheet.Cells[Specs.Btn_Visualize.Item1, Specs.Btn_Visualize.Item2];
                this.xlButton_Cancel = xlSheet.Cells[Specs.Btn_Cancel.Item1, Specs.Btn_Cancel.Item2];

                //LINK
                this.LinkToOrigin = new Data.Link(link.ToString());

                //Build the CorrelMatrix
                Excel.Range matrixRange = xlSheet.Range[xlMatrixCell.Offset[1, 0], xlMatrixCell.End[Excel.XlDirection.xlToRight].End[Excel.XlDirection.xlDown]];
                this.CorrelMatrix = Data.CorrelationMatrix.ConstructForConversion(pairs.GetCorrelationMatrix_Formulas(this), ids, fields, header);
                this.Header = header.ToString();
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
            //    this.xlHeaderCell.Value = this.CorrelString.Value;
            //}

            protected string GetSubIdString(IHasSubs est, int subIndex)
            {
                return ((IHasDurationCorrelations)est).SubEstimates[subIndex].uID.ID;
            }

            public override void PrintToSheet()  //expanding from string
            {
                CostSheet costSheet = CostSheet.ConstructFromXlCostSheet(this.LinkToOrigin.LinkSource.Worksheet);
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

                int subCount = parentEstimate.SubEstimates.Count();
                this.CorrelMatrix.PrintToSheet(xlMatrixCell);                                   //Print the matrix
                this.LinkToOrigin.PrintToSheet(xlLinkCell);                                     //Print the link
                this.xlHeaderCell.Value = this.Header;

                Excel.Range xlDistRange = xlDistCell.Resize[subCount, 1];
                object[,] xlDistValues = new object[subCount, 1];
                Excel.Range xlSubIdRange = xlSubIdCell.Resize[subCount, 1];
                object[,] xlSubIdValues = new object[subCount, 1];
                Excel.Range xlPairsRange = xlPairsCell.Resize[subCount - 1, 2];

                for (int subIndex = 0; subIndex < subCount; subIndex++)      //Print the Distribution strings
                {
                    xlDistValues[subIndex, 0] = ((Estimate_Item)parentEstimate).GetDistributionString(subIndex);
                    xlSubIdValues[subIndex, 0] = GetSubIdString(parentEstimate, subIndex);
                }
                xlDistRange.Value = xlDistValues;
                xlSubIdRange.Value = xlSubIdValues;
                xlSubIdRange.NumberFormat = "\"ID\";;;\"ID\"";
                xlPairsRange.Value = this.PairSpec.GetValuesString_Split();

                //Print Headers
                this.xlPairsCell.Offset[-1, 0].Value = "Off-diagonal Values";
                this.xlPairsCell.Offset[-1, 1].Value = "Linear reduction";
                this.xlSubIdCell.Offset[-1, 0].Value = "Unique ID";
                this.xlDistCell.Offset[-1, 0].Value = "Distribution";

                //I don't think these are necessary
                //this.Specs.PrintMatrixCoords(xlSheet);                                          //Print the matrix start coords
                //this.Specs.PrintLinkCoords(xlSheet);                                            //Print the link coords
                //this.Specs.PrintIdCoords(xlSheet);                                              //Print the ID coords
                //this.Specs.PrintDistCoords(xlSheet);                                            //Print the Distribution coords
                //this.PrintMatrixEndCoords(xlSheet);                                             //Print the matrix end coords

                AddUserControls();
                FormatSheet();
            }

            private void AddUserControls()
            {
                Vsto.Worksheet vstoSheet = Globals.Factory.GetVstoObject(this.xlSheet);
                System.Windows.Forms.Button btn_ConvertToDM = new System.Windows.Forms.Button();
                btn_ConvertToDM.Text = "Convert to Matrix Specification";
                btn_ConvertToDM.Click += ConversionFormClicked;
                vstoSheet.Controls.AddControl(btn_ConvertToDM, this.xlButton_ConvertCorrel.Resize[2,3], "ConvertToDM");

                //COLLAPSE
                System.Windows.Forms.Button btn_CollapseCorrelation = new System.Windows.Forms.Button();
                btn_CollapseCorrelation.Text = "Save Correlation";
                btn_CollapseCorrelation.Click += CollapseCorrelationClicked;
                vstoSheet.Controls.AddControl(btn_CollapseCorrelation, this.xlButton_CollapseCorrel.Resize[2, 3], "CollapseToCostSheet");

                //VISUALIZE
                System.Windows.Forms.Button btn_VisualizeCorrelation = new System.Windows.Forms.Button();
                btn_VisualizeCorrelation.Text = "Visualize";
                btn_VisualizeCorrelation.Click += VisualizeCorrelationClicked;
                vstoSheet.Controls.AddControl(btn_VisualizeCorrelation, this.xlButton_Visualize.Resize[2, 3], "VisualizeCorrelation");

                //CANCEL
                System.Windows.Forms.Button btn_Cancel = new System.Windows.Forms.Button();
                btn_Cancel.Text = "Cancel Changes";
                btn_Cancel.Click += CancelChangesClicked;
                vstoSheet.Controls.AddControl(btn_Cancel, this.xlButton_Cancel.Resize[2, 3], "CancelCorrelationChanges");

            }

            public override void FormatSheet()
            {
                Excel.Range matrixStart = this.xlMatrixCell.Offset[1, 0];
                Excel.Range matrixRange = matrixStart.Resize[this.CorrelMatrix.Fields.Length, this.CorrelMatrix.Fields.Length];
                Excel.Range diagonal = matrixRange.Cells[1,1];
                for(int i = 2; i <= matrixRange.Columns.Count; i++)
                {
                    diagonal = ThisAddIn.MyApp.Union(diagonal, matrixRange.Cells[i,i]);
                }
                Excel.Range pairwiseRange = this.xlPairsCell.Resize[matrixRange.Columns.Count - 1, 2];

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

                diagonal.Interior.Color = System.Drawing.Color.FromArgb(0, 0, 0);
                diagonal.Font.Color = System.Drawing.Color.FromArgb(255, 255, 255);

                matrixRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                matrixRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                pairwiseRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                pairwiseRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            }

            public override void ConvertCorrelation(bool PreserveOffDiagonal=false)
            {
                /*
                 * This method needs to construct a _DM type using the information on the _DP type.
                 * This includes fitting the pairs to a matrix.
                 * Need the fields, matrix, IDs, Link, Header
                 */
                object[,] matrix = this.CorrelMatrix.GetMatrix_Formulas();
                object[] ids = this.GetIDs();   //This isn't returning anything
                object[] fields = this.CorrelMatrix.Fields;
                object header = this.xlHeaderCell.Value;
                object link = this.xlLinkCell.Value;
                Sheets.CorrelationSheet_DM newSheet = new Sheets.CorrelationSheet_DM(matrix, ids, fields, header, link, this.xlSheet);
                newSheet.PrintToSheet();
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
