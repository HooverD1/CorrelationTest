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
        public class CorrelationSheet_CP : CorrelationSheet, IPairwiseSpec
        {
            public PairSpecification PairSpec { get; set; }
            public Data.CorrelationString_CP CorrelString { get; set; }
            public Excel.Range xlButton_ConvertCorrel { get; set; }

            //EXPAND
            public CorrelationSheet_CP(IHasCostCorrelations ParentItem)        //bring in the coordinates and set up the ranges once they exist
            {
                this.CorrelString = (Data.CorrelationString_CP)ParentItem.CostCorrelationString;
                this.Specs = new Data.CorrelSheetSpecs(SheetType.Correlation_CP);
                this.xlSheet = GetXlSheet();
                this.LinkToOrigin = new Data.Link(ParentItem.xlCorrelCell_Cost);
                this.xlLinkCell = xlSheet.Cells[Specs.LinkCoords.Item1, Specs.LinkCoords.Item2];
                this.xlHeaderCell = xlSheet.Cells[Specs.StringCoords.Item1, Specs.StringCoords.Item2];
                //this.xlIDCell = xlSheet.Cells[Specs.IdCoords.Item1, Specs.IdCoords.Item2];
                this.xlSubIdCell = xlSheet.Cells[Specs.SubIdCoords.Item1, Specs.SubIdCoords.Item2];
                this.xlDistCell = xlSheet.Cells[Specs.DistributionCoords.Item1, Specs.DistributionCoords.Item2];
                this.xlMatrixCell = xlSheet.Cells[Specs.MatrixCoords.Item1, Specs.MatrixCoords.Item2];
                this.xlPairsCell = xlSheet.Cells[Specs.PairsCoords.Item1, Specs.PairsCoords.Item2];
                this.xlButton_ConvertCorrel = xlSheet.Cells[Specs.Btn_ConvertCoords.Item1, Specs.Btn_ConvertCoords.Item2];
                this.xlButton_CollapseCorrel = xlSheet.Cells[Specs.Btn_Collapse.Item1, Specs.Btn_Collapse.Item2];
                this.xlButton_Visualize = xlSheet.Cells[Specs.Btn_Visualize.Item1, Specs.Btn_Visualize.Item2];
                this.xlButton_Cancel = xlSheet.Cells[Specs.Btn_Cancel.Item1, Specs.Btn_Cancel.Item2];
                CorrelMatrix = Data.CorrelationMatrix.ConstructFromParentItem(ParentItem, SheetType.Correlation_CP, this);
                this.Header = CorrelString.GetHeader();
                this.PairSpec = CorrelString.GetPairwise();

                //Should these be in PrintSheet()?

            }

            //COLLAPSE METHOD
            public CorrelationSheet_CP() //build from the xlsheet to get the string
            {
                this.xlSheet = GetXlSheet();
                this.Specs = new Data.CorrelSheetSpecs(SheetType.Correlation_CP);
                //Set up the xlCells
                this.xlLinkCell = xlSheet.Cells[Specs.LinkCoords.Item1, Specs.LinkCoords.Item2];
                this.xlHeaderCell = xlSheet.Cells[Specs.StringCoords.Item1, Specs.StringCoords.Item2];
                this.xlDistCell = xlSheet.Cells[Specs.DistributionCoords.Item1, Specs.DistributionCoords.Item2];
                this.xlSubIdCell = xlSheet.Cells[Specs.SubIdCoords.Item1, Specs.SubIdCoords.Item2];
                this.xlMatrixCell = xlSheet.Cells[Specs.MatrixCoords.Item1, Specs.MatrixCoords.Item2];
                this.xlPairsCell = xlSheet.Cells[Specs.PairsCoords.Item1, Specs.PairsCoords.Item2];
                this.xlButton_ConvertCorrel = xlSheet.Cells[Specs.Btn_ConvertCoords.Item1, Specs.Btn_ConvertCoords.Item2];
                this.LinkToOrigin = new Data.Link(Convert.ToString(xlLinkCell.Value));
                //
                //Build the CorrelMatrix
                object[] ids = Data.CorrelationString.GetIDsFromString(xlHeaderCell.Value);
                object[,] fieldsValues = xlSheet.Range[xlMatrixCell, xlMatrixCell.End[Excel.XlDirection.xlToRight]].Value;
                fieldsValues = ExtensionMethods.ReIndexArray(fieldsValues);
                object[] fields = ExtensionMethods.ToJaggedArray(fieldsValues)[0];
                int size = Data.CorrelationString.GetNumberOfInputsFromCorrelStringValue(xlHeaderCell.Value);
                Excel.Range matrixRange = xlMatrixCell.Offset[1, 0].Resize[size,size]; //xlSheet.Range[xlMatrixCell.Offset[1, 0], xlMatrixCell.End[Excel.XlDirection.xlDown].End[Excel.XlDirection.xlToRight]];
                object[,] matrix = matrixRange.Value;
                this.CorrelMatrix = Data.CorrelationMatrix.ConstructFromCorrelationSheet(this);
                this.Header = Convert.ToString(xlHeaderCell.Value);

                PairSpecification pairs = PairSpecification.ConstructFromRange(xlPairsCell, fields.Length - 1);
            }

            //CONVERT
            public CorrelationSheet_CP(PairSpecification pairs, object[] ids, object[] fields, object header, object link, Excel.Worksheet replaceXlSheet)
            {
                ThisAddIn.MyApp.DisplayAlerts = false;
                replaceXlSheet.Delete();
                ThisAddIn.MyApp.DisplayAlerts = true;
                this.xlSheet = GetXlSheet();
                this.Specs = new Data.CorrelSheetSpecs(SheetType.Correlation_CP);
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

                this.LinkToOrigin = new Data.Link(link.ToString());
                this.Header = header.ToString();

                //Build the CorrelMatrix
                Excel.Range matrixRange = xlSheet.Range[xlMatrixCell.Offset[1, 0], xlMatrixCell.End[Excel.XlDirection.xlToRight].End[Excel.XlDirection.xlDown]];
                this.CorrelMatrix = Data.CorrelationMatrix.ConstructForConversion(pairs.GetCorrelationMatrix_Formulas(this), ids, fields, header);
                this.PairSpec = pairs;
            }

            public override void FormatSheet()
            {
                int size = this.CorrelMatrix.Fields.Length;
                Excel.Range matrixStart = this.xlMatrixCell.Offset[1, 0];
                Excel.Range matrixRange = matrixStart.Resize[size, size];
                Excel.Range xlPairsRange = this.xlPairsCell.Resize[size - 1, 2];
                Excel.Range diagonal = matrixRange.Cells[1, 1];
                for (int i = 2; i <= size; i++)
                    diagonal = ThisAddIn.MyApp.Union(diagonal, matrixRange.Cells[i, i]);
                foreach (Excel.Range row in matrixRange.Rows)
                {
                    row.Interior.Color = System.Drawing.Color.FromArgb(225, 225, 225);      //Dying right here if you try to do it all at once?
                    row.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    row.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                }

                xlPairsRange.Interior.Color = System.Drawing.Color.FromArgb(255, 255, 190);
                xlPairsRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                xlPairsRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                diagonal.Interior.Color = System.Drawing.Color.FromArgb(0, 0, 0);
                diagonal.Font.Color = System.Drawing.Color.FromArgb(255, 255, 255);

                this.xlSheet.Cells.Columns.AutoFit();
                this.xlSheet.Columns[1].ColumnWidth = 25;
            }

            protected override Excel.Worksheet GetXlSheet(bool CreateNew = true)
            {
                var xlCorrelSheets = from Excel.Worksheet sheet in ThisAddIn.MyApp.Worksheets
                                     where Convert.ToString(sheet.Cells[1, 1].value) == "$CORRELATION_CP"
                                     select sheet;
                if (xlCorrelSheets.Any())
                    xlSheet = xlCorrelSheets.First();
                else if (CreateNew)
                    xlSheet = CreateXLCorrelSheet("_CP");
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

            protected string GetSubIdString(IHasSubs est, int subIndex)
            {
                return ((IHasCostCorrelations)est).SubEstimates[subIndex].uID.ID;
            }


            public override void PrintToSheet()  //expanding from string
            {
                CostSheet costSheet = CostSheet.ConstructFromXlCostSheet(this.LinkToOrigin.LinkSource.Worksheet);
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
                int subCount = parentEstimate.SubEstimates.Count();
                this.CorrelMatrix.PrintToSheet(xlMatrixCell);                                   //Print the matrix
                this.LinkToOrigin.PrintToSheet(xlLinkCell);                                     //Print the link

                this.xlHeaderCell.Value = this.Header;
                this.xlHeaderCell.NumberFormat = "\"CORREL\";;;\"CORREL\"";
                this.xlHeaderCell.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                this.xlPairsCell.Resize[subCount - 1, 2].Value = this.PairSpec.GetValuesString_Split();

                Excel.Range xlDistRange = xlDistCell.Resize[subCount, 1];
                object[,] xlDistValues = new object[subCount, subCount];

                Excel.Range xlSubIdRange = xlSubIdCell.Resize[subCount, 1];
                object[,] xlSubIdValues = new object[subCount, subCount];

                for (int subIndex = 0; subIndex < subCount; subIndex++)      //Load object arrays for printing
                {
                    xlDistValues[subIndex, 0] = ((Estimate_Item)parentEstimate).GetDistributionString(subIndex);
                    xlSubIdValues[subIndex, 0] = GetSubIdString(parentEstimate, subIndex);
                }
                xlDistRange.Value = xlDistValues;                           //Print object arrays
                xlDistRange.NumberFormat = "\"DIST\";;;\"DIST\"";
                xlSubIdRange.Value = xlSubIdValues;
                xlSubIdRange.NumberFormat = "\"ID\";;;\"ID\"";

                //Print column headers
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

                
                FormatSheet();

                AddUserControls();

                
            }

            private void AddUserControls()
            {
                Vsto.Worksheet vstoSheet = Globals.Factory.GetVstoObject(this.xlSheet);

                //CONVERT
                System.Windows.Forms.Button btn_ConvertToCM = new System.Windows.Forms.Button();
                btn_ConvertToCM.Text = "Convert to Matrix Specification";
                btn_ConvertToCM.Click += ConversionFormClicked;
                vstoSheet.Controls.AddControl(btn_ConvertToCM, this.xlButton_ConvertCorrel.Resize[2, 1], "ConvertToCM");

                //COLLAPSE
                System.Windows.Forms.Button btn_CollapseCorrelation = new System.Windows.Forms.Button();
                btn_CollapseCorrelation.Text = "Save Correlation";
                btn_CollapseCorrelation.Click += CollapseCorrelationClicked;
                vstoSheet.Controls.AddControl(btn_CollapseCorrelation, this.xlButton_CollapseCorrel.Resize[2, 1], "CollapseToCostSheet");

                //VISUALIZE
                System.Windows.Forms.Button btn_VisualizeCorrelation = new System.Windows.Forms.Button();
                btn_VisualizeCorrelation.Text = "Visualize";
                btn_VisualizeCorrelation.Click += VisualizeCorrelationClicked;
                vstoSheet.Controls.AddControl(btn_VisualizeCorrelation, this.xlButton_Visualize.Resize[2, 1], "VisualizeCorrelation");

                //CANCEL
                System.Windows.Forms.Button btn_Cancel = new System.Windows.Forms.Button();
                btn_Cancel.Text = "Cancel Changes";
                btn_Cancel.Click += CancelChangesClicked;
                vstoSheet.Controls.AddControl(btn_Cancel, this.xlButton_Cancel.Resize[2, 1], "CancelCorrelationChanges");

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
                 * This method needs to construct a _CM type using the information on the _CP type.
                 * This includes fitting the pairs to a matrix.
                 * Need the fields, matrix, IDs, Link, Header
                 */
                
                object[,] matrix = this.CorrelMatrix.GetMatrix_Formulas();
                object[] ids = this.GetIDs();   //This isn't returning anything
                object[] fields = this.CorrelMatrix.Fields;
                object header = this.xlHeaderCell.Value;
                object link = this.xlLinkCell.Value;
                Sheets.CorrelationSheet_CM newSheet = new Sheets.CorrelationSheet_CM(matrix, ids, fields, header, link, this.xlSheet);
                newSheet.PrintToSheet();
            }
        }
    }
}
