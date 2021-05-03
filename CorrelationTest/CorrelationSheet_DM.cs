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
        public class CorrelationSheet_DM : CorrelationSheet, IMatrixSpec
        {
            public Data.CorrelationString_DM CorrelString { get; set; }
            public Excel.Range xlButton_ConvertCorrel { get; set; }

            //EXPAND
            public CorrelationSheet_DM(IHasDurationCorrelations ParentItem)        //bring in the coordinates and set up the ranges once they exist
            {   //Build from the correlString to get the xlSheet
                this.CorrelString = (Data.CorrelationString_DM)ParentItem.DurationCorrelationString;
                this.Specs = new Data.CorrelSheetSpecs(SheetType.Correlation_DM);
                this.xlSheet = GetXlSheet();
                this.LinkToOrigin = new Data.Link(ParentItem.xlCorrelCell_Duration);
                this.xlLinkCell = xlSheet.Cells[Specs.LinkCoords.Item1, Specs.LinkCoords.Item2];
                this.xlHeaderCell = xlSheet.Cells[Specs.StringCoords.Item1, Specs.StringCoords.Item2];
                //this.xlIDCell = xlSheet.Cells[Specs.IdCoords.Item1, Specs.IdCoords.Item2];      //Is this junk?
                this.xlSubIdCell = xlSheet.Cells[Specs.SubIdCoords.Item1, Specs.SubIdCoords.Item2];
                this.xlDistCell = xlSheet.Cells[Specs.DistributionCoords.Item1, Specs.DistributionCoords.Item2];
                this.xlMatrixCell = xlSheet.Cells[Specs.MatrixCoords.Item1, Specs.MatrixCoords.Item2];
                this.xlButton_ConvertCorrel = xlSheet.Cells[Specs.Btn_ConvertCoords.Item1, Specs.Btn_ConvertCoords.Item2];
                this.xlButton_CollapseCorrel = xlSheet.Cells[Specs.Btn_Collapse.Item1, Specs.Btn_Collapse.Item2];
                this.xlButton_Visualize = xlSheet.Cells[Specs.Btn_Visualize.Item1, Specs.Btn_Visualize.Item2];
                this.xlButton_Cancel = xlSheet.Cells[Specs.Btn_Cancel.Item1, Specs.Btn_Cancel.Item2];

                CorrelMatrix = Data.CorrelationMatrix.ConstructFromParentItem(ParentItem, SheetType.Correlation_DM, this);
                this.Header = CorrelString.GetHeader();
            }
            //COLLAPSE
            public CorrelationSheet_DM() //build from the xlsheet to get the string
            {
                //Need a link
                this.xlSheet = GetXlSheet();
                this.Specs = new Data.CorrelSheetSpecs(SheetType.Correlation_DM);
                //Set up the xlCells
                this.xlLinkCell = xlSheet.Cells[Specs.LinkCoords.Item1, Specs.LinkCoords.Item2];
                this.xlHeaderCell = xlSheet.Cells[Specs.StringCoords.Item1, Specs.StringCoords.Item2];
                //this.xlIDCell = xlSheet.Cells[Specs.IdCoords.Item1, Specs.IdCoords.Item2];
                this.xlDistCell = xlSheet.Cells[Specs.DistributionCoords.Item1, Specs.DistributionCoords.Item2];
                this.xlSubIdCell = xlSheet.Cells[Specs.SubIdCoords.Item1, Specs.SubIdCoords.Item2];
                this.xlMatrixCell = xlSheet.Cells[Specs.MatrixCoords.Item1, Specs.MatrixCoords.Item2];
                this.xlPairsCell = xlSheet.Cells[Specs.PairsCoords.Item1, Specs.PairsCoords.Item2];
                this.xlButton_ConvertCorrel = xlSheet.Cells[Specs.Btn_ConvertCoords.Item1, Specs.Btn_ConvertCoords.Item2];
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
                string parent_id = Data.CorrelationString.GetParentIDFromCorrelStringValue(xlHeaderCell.Value);
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
                this.xlHeaderCell = this.xlSheet.Cells[this.Specs.StringCoords.Item1, this.Specs.StringCoords.Item2];
                //this.xlIDCell = this.xlSheet.Cells[this.Specs.IdCoords.Item1, this.Specs.IdCoords.Item2];
                this.xlDistCell = this.xlSheet.Cells[this.Specs.DistributionCoords.Item1, this.Specs.DistributionCoords.Item2];
                this.xlSubIdCell = this.xlSheet.Cells[this.Specs.SubIdCoords.Item1, this.Specs.SubIdCoords.Item2];
                this.xlMatrixCell = this.xlSheet.Cells[this.Specs.MatrixCoords.Item1, this.Specs.MatrixCoords.Item2];
                this.xlButton_ConvertCorrel = xlSheet.Cells[Specs.Btn_ConvertCoords.Item1, Specs.Btn_ConvertCoords.Item2];
                this.xlButton_CollapseCorrel = xlSheet.Cells[Specs.Btn_Collapse.Item1, Specs.Btn_Collapse.Item2];
                this.xlButton_Visualize = xlSheet.Cells[Specs.Btn_Visualize.Item1, Specs.Btn_Visualize.Item2];
                this.xlButton_Cancel = xlSheet.Cells[Specs.Btn_Cancel.Item1, Specs.Btn_Cancel.Item2];

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
                this.Header = header.ToString();
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
                int subCount = parentEstimate.SubEstimates.Count();
                this.CorrelMatrix.PrintToSheet(xlMatrixCell);                                   //Print the matrix
                this.LinkToOrigin.PrintToSheet(xlLinkCell);                                     //Print the link

                this.xlHeaderCell.Value = this.Header;
                this.xlHeaderCell.NumberFormat = "\"CORREL\";;;\"CORREL\"";
                this.xlHeaderCell.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                object[,] xlDistValues = new object[subCount, 1];
                Excel.Range xlDistRange = xlDistCell.Resize[subCount, 1];

                object[,] xlSubIdValues = new object[subCount, 1];
                Excel.Range xlSubIdRange = xlSubIdCell.Resize[subCount, 1];

                for (int subIndex = 0; subIndex < subCount; subIndex++)      //Print the Distribution strings
                {
                    xlDistValues[subIndex, 0] = ((Estimate_Item)parentEstimate).GetDistributionString(subIndex);
                    xlSubIdValues[subIndex, 0] = GetSubIdString(parentEstimate, subIndex);
                }

                xlDistRange.Value = xlDistValues;
                xlDistRange.NumberFormat = "\"DIST\";;;\"DIST\"";
                xlSubIdRange.Value = xlSubIdValues;
                xlSubIdRange.NumberFormat = "\"ID\";;;\"ID\"";

                this.xlSubIdCell.Offset[-1, 0].Value = "Unique ID";
                this.xlDistCell.Offset[-1, 0].Value = "Distribution";

                AddUserControls();
                FormatSheet();
            }

            private void AddUserControls()
            {
                Vsto.Worksheet vstoSheet = Globals.Factory.GetVstoObject(this.xlSheet);
                System.Windows.Forms.Button btn_ConvertToDP = new System.Windows.Forms.Button();
                btn_ConvertToDP.Text = "Convert to Pairwise Specification";
                btn_ConvertToDP.Click += ConversionFormClicked;
                vstoSheet.Controls.AddControl(btn_ConvertToDP, this.xlButton_ConvertCorrel.Resize[2, 1], "ConvertToDP");

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

                this.xlSheet.Cells.Columns.AutoFit();
                this.xlSheet.Columns[1].ColumnWidth = 25;
            }

            public override void ConvertCorrelation(bool PreserveOffDiagonal=false) //Convert DP --> DM (fit matrix)
            {
                /*
                 * This method needs to construct a _DM type using the information on the _DP type.
                 * This includes fitting the pairs to a matrix.
                 * Need the fields, matrix, IDs, Link, Header
                 */
                var pairs = PairSpecification.ConstructByFittingMatrix(this.CorrelMatrix.GetMatrix_Values(), PreserveOffDiagonal);
                object[] ids = this.GetIDs();
                object[] fields = this.GetFields();
                object header = this.xlHeaderCell.Value;
                object link = this.xlLinkCell.Value;
                CorrelationSheet_DP convertedSheet = new CorrelationSheet_DP(pairs, ids, fields, header, link, this.xlSheet);
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
