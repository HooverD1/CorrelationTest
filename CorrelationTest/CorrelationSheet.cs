using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace CorrelationTest
{
    namespace Sheets
    {
        public class CorrelationSheet : Sheet
        {
            public readonly static Tuple<int, int> param_RowLink = new Tuple<int, int>(1, 2);    //where to find the row param for the link
            public readonly static Tuple<int, int> param_ColLink = new Tuple<int, int>(1, 3);    //where to find the col param for the link
            public readonly static Tuple<int, int> param_RowMatrix = new Tuple<int, int>(1, 4);    //where to find the row param for the matrix
            public readonly static Tuple<int, int> param_ColMatrix = new Tuple<int, int>(1, 5);    //where to find the col param for the matrix
            public readonly static Tuple<int, int> param_RowMatrix2 = new Tuple<int, int>(1, 6);    //where to find the last row param for the matrix
            public readonly static Tuple<int, int> param_ColMatrix2 = new Tuple<int, int>(1, 7);    //where to find the last col param for the matrix
            public readonly static Tuple<int, int> param_RowID = new Tuple<int, int>(1, 8);    //where to find the last col param for the ID
            public readonly static Tuple<int, int> param_ColID = new Tuple<int, int>(1, 9);    //where to find the last col param for the ID
            public readonly static Tuple<int, int> param_RowDist = new Tuple<int, int>(1, 10);    //where to find the last col param for the Distribution
            public readonly static Tuple<int, int> param_ColDist = new Tuple<int, int>(1, 11);    //where to find the last col param for the Distribution

            //public Data.CorrelationString CorrelString { get; set; }
            public Data.CorrelationMatrix CorrelMatrix { get; set; }
            public Excel.Range xlMatrixCell { get; set; }
            public Data.Link LinkToOrigin { get; set; }
            public Excel.Range xlLinkCell { get; set; }
            //public Excel.Range xlIDCell { get; set; }
            public Excel.Range xlHeaderCell { get; set; }
            public Excel.Range xlDistCell { get; set; }
            public Excel.Range xlSubIdCell { get; set; }
            public Data.CorrelSheetSpecs Specs { get; set; }
            public Excel.Range xlPairsCell { get; set; }

            public Excel.Range xlButton_CollapseCorrel { get; set; }
            public Excel.Range xlButton_Visualize { get; set; }
            public Excel.Range xlButton_Cancel { get; set; }

            public string Header { get; set; }

            //public CorrelationSheet(Data.CorrelationString_CM correlString, Excel.Range launchedFrom) : this(correlString, launchedFrom, new Data.CorrelSheetSpecs()) { }       //default locations


            protected virtual Excel.Worksheet CreateXLCorrelSheet(string postfix) { throw new Exception("Failed override"); }
            protected virtual Excel.Worksheet GetXlSheet(bool CreateNew = true) { throw new Exception("Failed override"); }
            public virtual void UpdateCorrelationString(string[] ids) { throw new Exception("Failed override"); }

            public virtual string[] GetIDs()
            {
                //This needs to pull the IDs range off the sheet, not use the correl string - which no longer appears on the sheet fully
                int numberOfSubs = Data.CorrelationString.GetNumberOfInputsFromCorrelStringValue(this.xlHeaderCell.Value);
                Excel.Range xlSubIdRange = xlSubIdCell.Resize[numberOfSubs, 1];
                string[] ids = new string[numberOfSubs];
                for (int i = 0; i < numberOfSubs; i++)
                    ids[i] = Convert.ToString(xlSubIdRange.Cells[i + 1, 1].value);
                return ids;
            }

            public virtual string[] GetFields()
            {
                int numberOfSubs = Data.CorrelationString.GetNumberOfInputsFromCorrelStringValue(this.xlHeaderCell.Value);
                Excel.Range xlSubIdRange = xlMatrixCell.Offset[1,-1].Resize[numberOfSubs, 1];
                string[] fields = new string[numberOfSubs];
                for (int i = 0; i < numberOfSubs; i++)
                    fields[i] = Convert.ToString(xlSubIdRange.Cells[i + 1, 1].value);
                return fields;
            }

            public virtual object[,] GetMatrix()
            {
                int numberOfSubs = Data.CorrelationString.GetNumberOfInputsFromCorrelStringValue(this.xlHeaderCell.Value);
                Excel.Range xlMatrixRange = xlMatrixCell.Offset[1, 0].Resize[numberOfSubs, numberOfSubs];
                return xlMatrixRange.Value;
            }

            public Data.CorrelationString CollapseToString(object[,] correlArray)
            {
                throw new NotImplementedException();
            }


            public double[,] SetCorrelArray(Data.CorrelationMatrix correlMatrix)
            {
                throw new NotImplementedException();
            }
            public string[] Get_xlFields()
            {
                Excel.Range endCell = this.xlMatrixCell.End[Excel.XlDirection.xlToRight];
                Excel.Range fieldRange = this.xlSheet.Range[this.xlMatrixCell, endCell];
                object[,] fieldRangeValues = ExtensionMethods.ReIndexArray(fieldRange.Value);
                object[][] jaggedRange = ExtensionMethods.ToJaggedArray(fieldRangeValues);
                string[] returnString = new string[jaggedRange[0].Length];
                for(int i = 0; i < jaggedRange[0].Length; i++)
                {
                    returnString[i] = Convert.ToString(jaggedRange[0][i]);
                }
                return returnString;
            }

            private Tuple<int, int> GetMatrixEndCoords(Excel.Range xlMatrixCell, int fieldCount)
            {
                return new Tuple<int, int>(xlMatrixCell.Row + fieldCount, xlMatrixCell.Column + fieldCount-1);
            }
            protected void PrintMatrixEndCoords(Excel.Worksheet xlCorrelSheet)
            {
                this.Specs.MatrixCoords_End = GetMatrixEndCoords(xlMatrixCell, this.CorrelMatrix.FieldCount);
                xlCorrelSheet.Cells[Sheets.CorrelationSheet.param_RowMatrix2.Item1, Sheets.CorrelationSheet.param_RowMatrix2.Item2].Value = this.Specs.MatrixCoords_End.Item1;  //prints the row coord
                xlCorrelSheet.Cells[Sheets.CorrelationSheet.param_ColMatrix2.Item1, Sheets.CorrelationSheet.param_ColMatrix2.Item2].Value = this.Specs.MatrixCoords_End.Item2;  //prints the col coord
            }

            public override void PrintToSheet() { throw new Exception("Failed override"); }

            protected string[] GetFieldsFromXlCorrelSheet()
            {
                Excel.Range fieldEndCell = this.xlMatrixCell.End[Excel.XlDirection.xlToRight];
                Excel.Range fieldRange = this.xlSheet.Range[xlMatrixCell, fieldEndCell];
                string[] fieldStrings = new string[fieldRange.Cells.Count];
                for (int i = 1; i < fieldRange.Cells.Count; i++)
                {
                    fieldStrings[i - 1] = Convert.ToString(fieldRange.Cells[1, i].value);
                }
                return fieldStrings;
            }

            protected virtual string GetDistributionString(IHasCorrelations est, int subIndex) { throw new Exception("Failed override"); }
            protected virtual string GetDistributionString(IHasCorrelations est) { throw new Exception("Failed override"); }

            public bool PaintMatrixErrors(Data.MatrixErrors[,] matrixErrors)        //return false if no matrix errors found
            {
                IEnumerable<Data.MatrixErrors> findErrors = from Data.MatrixErrors me in matrixErrors
                                                            where me != Data.MatrixErrors.None
                                                            select me;
                if (findErrors.Any())
                {
                    for (int row = 0; row < matrixErrors.GetLength(0); row++)
                    {
                        for (int col = row; col < matrixErrors.GetLength(1); col++)
                        {
                            if (matrixErrors[row, col] == Data.MatrixErrors.None)
                            {
                                this.xlMatrixCell.Offset[row + 1, col].ClearComments();
                                if(row != col)
                                    this.xlMatrixCell.Offset[row + 1, col].Interior.ColorIndex = 8;
                            }
                            else if (matrixErrors[row, col] == Data.MatrixErrors.AboveUpperBound)
                            {
                                this.xlMatrixCell.Offset[row + 1, col].Interior.Color = Excel.XlRgbColor.rgbRed;
                                this.xlMatrixCell.Offset[row + 1, col].ClearComments();
                                this.xlMatrixCell.Offset[row + 1, col].AddComment("Above Upper Bound");
                            }
                            else if (matrixErrors[row, col] == Data.MatrixErrors.BelowLowerBound)
                            {
                                this.xlMatrixCell.Offset[row + 1, col].Interior.Color = Excel.XlRgbColor.rgbRed;
                                this.xlMatrixCell.Offset[row + 1, col].ClearComments();
                                this.xlMatrixCell.Offset[row + 1, col].AddComment("Below Lower Bound");
                            }
                            else if (matrixErrors[row, col] == Data.MatrixErrors.MisplacedValue)
                            {
                                this.xlMatrixCell.Offset[row + 1, col].Interior.Color = Excel.XlRgbColor.rgbRed;
                                this.xlMatrixCell.Offset[row + 1, col].ClearComments();
                                this.xlMatrixCell.Offset[row + 1, col].AddComment("Only fill upper triangular portion");
                            }
                        }
                    }
                    return true;
                }
                else   //No errors found -- don't paint, return false
                {
                    return false;
                }
            }

            private static Excel.Worksheet GetCorrelationSheet()
            {
                List<Excel.Worksheet> xlCorrelSheets = new List<Excel.Worksheet>();
                foreach (Excel.Worksheet sht in ThisAddIn.MyApp.Worksheets)
                {
                    SheetType sht_type = ExtensionMethods.GetSheetType(sht);
                    if (sht_type == SheetType.Correlation_DM ||
                        sht_type == SheetType.Correlation_DP ||
                        sht_type == SheetType.Correlation_CP ||
                        sht_type == SheetType.Correlation_CM ||
                        sht_type == SheetType.Correlation_PM ||
                        sht_type == SheetType.Correlation_PP)
                    {
                        xlCorrelSheets.Add(sht);
                    }
                }
                if (xlCorrelSheets.Count() == 1)
                    return xlCorrelSheets.First();
                else if (xlCorrelSheets.Count() > 1)
                    throw new Exception("Multiple correlation sheets");
                else if (!xlCorrelSheets.Any())
                    throw new Exception("No correlation sheets");
                else
                    throw new Exception("Unknown error finding correlation sheet");
            }

            public static CorrelationSheet ConstructFromXlCorrelationSheet()
            {
                Excel.Worksheet xlCorrelationSheet = GetCorrelationSheet();
                SheetType sheet_type = ExtensionMethods.GetSheetType(xlCorrelationSheet);
                //Make the sheetid cell in 1,1 list the type of correlation
                CorrelationSheet newSheet;
                switch (sheet_type)
                {
                    case SheetType.Correlation_DP:
                        newSheet = new CorrelationSheet_DP();
                        break;
                    case SheetType.Correlation_DM:
                        newSheet = new CorrelationSheet_DM();
                        break;
                    case SheetType.Correlation_CP:
                        newSheet = new CorrelationSheet_CP();
                        break;
                    case SheetType.Correlation_CM:
                        newSheet = new CorrelationSheet_CM();
                        break;
                    case SheetType.Correlation_PP:
                        newSheet = new CorrelationSheet_PP();
                        break;
                    default:
                        throw new Exception("Not a valid Correlation Sheet type");
                    
                }
                //Why aren't these being done in the constructors...?
                newSheet.xlSheet = xlCorrelationSheet;
                //newSheet.CorrelString = Data.CorrelationString.ConstructFromCorrelationSheet(newSheet);
                return newSheet;
            }

            //EXPAND
            public static CorrelationSheet ConstructFromParentItem(IHasSubs ParentItem, SheetType CorrelType)
            {
                //find if it's cost, phasing, duration -- pass the selection?
                //Cast the parent item
                //Pick up its sheet type off the correlstring on the parent item
                //Remove correltype parameter
                CorrelationSheet returnSheet;
                switch (CorrelType)
                {
                    //These need to be sending the parent and the correltype, no?
                    case SheetType.Correlation_CM:
                        returnSheet = new CorrelationSheet_CM((IHasCostCorrelations)ParentItem);
                        break;
                    case SheetType.Correlation_CP:
                        returnSheet = new CorrelationSheet_CP((IHasCostCorrelations)ParentItem);
                        break;
                    case SheetType.Correlation_PP:
                        returnSheet = new CorrelationSheet_PP((IHasPhasingCorrelations)ParentItem);
                        break;
                    case SheetType.Correlation_DM:
                        returnSheet = new CorrelationSheet_DM((IHasDurationCorrelations)ParentItem);
                        break;
                    case SheetType.Correlation_DP:
                        returnSheet = new CorrelationSheet_DP((IHasDurationCorrelations)ParentItem);
                        break;
                    default:
                        throw new Exception("Unknown correlation type");
                }
                returnSheet.CorrelMatrix.ContainingSheet = returnSheet;
                return returnSheet;
            }

            

            

            private Data.CorrelStringType GetCorrelType(string correlStringValue)
            {
                switch (correlStringValue)
                {
                    case "CP":
                        return Data.CorrelStringType.CostPair;
                    case "CM":
                        return Data.CorrelStringType.CostMatrix;
                    case "PM":
                        return Data.CorrelStringType.PhasingMatrix;
                    case "PP":
                        return Data.CorrelStringType.PhasingPair;
                    case "DP":
                        return Data.CorrelStringType.DurationPair;
                    case "DM":
                        return Data.CorrelStringType.DurationMatrix;
                    default:
                        throw new Exception("Malformed ID");
                }
            }

            public static void CollapseToSheet()    //grab the xlSheet matrix, build the correlString from it, place it at the origin, delete the xlSheet
            {
                ExtensionMethods.TurnOffUpdating();
                CorrelationSheet correlSheet = CorrelationSheet.ConstructFromXlCorrelationSheet();
                //CorrelationType cType = ExtensionMethods.GetCorrelationTypeFromLink(correlSheet.LinkToOrigin.LinkSource);
                if (correlSheet == null)
                    return;

                //Validate matrix checks
                //Validate link source ID
                //Validate that the linkSource still has an ID match. If so, .PrintToSheet ... Otherwise, search for the ID and throw a warning ... if no ID can be found, throw an error and don't delete the sheet
                CostSheet originSheet = CostSheet.ConstructFromXlCostSheet(correlSheet.LinkToOrigin.LinkSource.Worksheet);
                string id_followLink = Convert.ToString(correlSheet.LinkToOrigin.LinkSource.EntireRow.Cells[1, originSheet.Specs.ID_Offset].value);
                string id_correlSheet = Data.CorrelationString.GetParentIDFromCorrelStringValue(correlSheet.xlHeaderCell.Value);
                if (CheckMatrixForErrors(correlSheet))
                {
                    //If matrix errors exist, kill the process.
                    ExtensionMethods.TurnOnUpdating();
                    return;
                }

                if (id_followLink.ToString() == id_correlSheet)
                {
                    Item sourceParent = (from Item item in originSheet.Items where item.uID.ID == id_correlSheet.ToString() select item).First();
                    if (correlSheet is Sheets.CorrelationSheet_CP)
                    {
                        Data.CorrelationString.ConstructFromCorrelationSheet(correlSheet).PrintToSheet((from ISub sub in ((IHasCostCorrelations)sourceParent).SubEstimates select sub.xlCorrelCell_Cost).ToArray());
                        //((Sheets.CorrelationSheet_CP)correlSheet).CorrelString.PrintToSheet((from ISub sub in ((IHasCostCorrelations)sourceParent).SubEstimates select sub.xlCorrelCell_Cost).ToArray());
                    }
                    else if(correlSheet is Sheets.CorrelationSheet_CM)
                    {
                        IEnumerable<Excel.Range> printEnumerable = from ISub sub in ((IHasCostCorrelations)sourceParent).SubEstimates select sub.xlCorrelCell_Cost;
                        Excel.Range[] printArray;
                        if (printEnumerable.Any())
                        {
                            printArray = printEnumerable.ToArray();
                            Data.CorrelationString.ConstructFromCorrelationSheet(correlSheet).PrintToSheet(printArray);
                        }
                        else
                        {
                            throw new Exception("No print area");
                        }
                        //((Sheets.CorrelationSheet_CM)correlSheet).CorrelString.PrintToSheet((from ISub sub in ((IHasCostCorrelations)sourceParent).SubEstimates select sub.xlCorrelCell_Cost).ToArray());
                    }
                    else if (correlSheet is Sheets.CorrelationSheet_DP)
                    {
                        Data.CorrelationString.ConstructFromCorrelationSheet(correlSheet).PrintToSheet((from ISub sub in ((IHasDurationCorrelations)sourceParent).SubEstimates select sub.xlCorrelCell_Duration).ToArray());
                        //((Sheets.CorrelationSheet_DP)correlSheet).CorrelString.PrintToSheet((from ISub sub in ((IHasDurationCorrelations)sourceParent).SubEstimates select sub.xlCorrelCell_Duration).ToArray());
                    }

                    else if(correlSheet is Sheets.CorrelationSheet_DM)
                    {
                        Data.CorrelationString.ConstructFromCorrelationSheet(correlSheet).PrintToSheet((from ISub sub in ((IHasDurationCorrelations)sourceParent).SubEstimates select sub.xlCorrelCell_Duration).ToArray());
                        //    ((Sheets.CorrelationSheet_DM)correlSheet).CorrelString.PrintToSheet((from ISub sub in ((IHasDurationCorrelations)sourceParent).SubEstimates select sub.xlCorrelCell_Duration).ToArray());
                    }
                    else if (correlSheet is Sheets.CorrelationSheet_PP)
                    {
                        Data.CorrelationString.ConstructFromCorrelationSheet(correlSheet).PrintToSheet(sourceParent.xlCorrelCell_Phasing);
                        //((Sheets.CorrelationSheet_PP)correlSheet).CorrelString.PrintToSheet(sourceParent.xlCorrelCell_Phasing);
                    }
                    else
                        throw new Exception("Unknown parent type");

                    ThisAddIn.MyApp.DisplayAlerts = false;
                    correlSheet.xlSheet.Delete();
                    correlSheet.LinkToOrigin.LinkSource.Worksheet.Activate();
                    correlSheet = null;
                    ThisAddIn.MyApp.DisplayAlerts = true;
                }                    
                else
                    MessageBox.Show("ID not found");
                
                ExtensionMethods.TurnOnUpdating();
            }

            public static bool CheckMatrixForErrors(CorrelationSheet correlSheet)
            {
                bool psd_errors = false;
                if (!correlSheet.CorrelMatrix.CheckForPSD())
                {
                    psd_errors = true;
                    DialogResult dialog_fixPSD = MessageBox.Show("Would you like to adjust this matrix to be positive semi-definite?", "Matrix is not PSD", MessageBoxButtons.YesNoCancel);
                    if (dialog_fixPSD == DialogResult.Cancel)
                    {
                        //Do nothing
                    }
                    else if (dialog_fixPSD == DialogResult.No)
                    {
                        //MessageBox.Show("Matrix cannot be saved.");
                    }
                    else if (dialog_fixPSD == DialogResult.Yes)
                    {
                        correlSheet.CorrelMatrix.FixMatrixForPSD();
                    }
                    //Launch form to correct PSD
                }
                bool trans_errors = correlSheet.PaintMatrixErrors(correlSheet.CorrelMatrix.CheckMatrixForTransitivity());         //This is stalling out a large matrix
                
                return psd_errors || trans_errors;
            }

            protected virtual object[,] GetDistributionParamStrings()
            {
                int lastRow = xlDistCell.End[Excel.XlDirection.xlDown].Row;
                Excel.Range distRange = xlDistCell.Resize[lastRow - xlDistCell.Row+1, 1];
                return distRange.Value;
            }

            public virtual void VisualizeCorrel()
            {
                //Select a correlation
                
                int selectionRow = ThisAddIn.MyApp.Selection.Row;       //dependency..
                int selectionCol = ThisAddIn.MyApp.Selection.Column;    //dependency..
                //Make sure the row and column are within the matrix range
                if (selectionRow < this.Specs.MatrixCoords.Item1 + 1 || selectionRow > this.CorrelMatrix.FieldCount + this.Specs.MatrixCoords.Item1)
                {
                    MessageBox.Show("Must select a matrix cell to visualize");
                    return;
                }
                if (selectionCol < this.Specs.MatrixCoords.Item2 || selectionCol > this.CorrelMatrix.FieldCount + this.Specs.MatrixCoords.Item2 - 1)
                {
                    MessageBox.Show("Must select a matrix cell to visualize");
                    return;
                }
                if(selectionRow - (this.Specs.MatrixCoords.Item1 + 1) == selectionCol - (this.Specs.MatrixCoords.Item2))
                {
                    //Along the diagonal - do nothing
                    return;
                } 
                //double coefficient = Convert.ToDouble(ThisAddIn.MyApp.Selection);
                //Find the distributions to use
                object[,] distParams = GetDistributionParamStrings();
                int distRow = selectionRow - xlMatrixCell.Row;
                int distCol = selectionCol - xlMatrixCell.Column;
                //Create distribution objects
                IEstimateDistribution d1 = Distribution.ConstructForVisualization(ThisAddIn.MyApp.Selection, this);
                //Need to load the row corresponding to the column selected (distCol)
                Excel.Range columnRowCell = this.xlDistCell.Offset[distCol, 0];
                IEstimateDistribution d2 = Distribution.ConstructForVisualization(columnRowCell, this);
                //Create the form
                if(d1==null || d2==null)
                {
                    //No distribution for one of these items -- cannot load the visual
                    MessageBox.Show("Selected correlation lacks distribution information.");
                    return;
                }
                CorrelationForm CorrelVisual = new CorrelationForm(d1, d2, 0);
                CorrelVisual.StartPosition = FormStartPosition.Manual;
                CorrelVisual.Location = new System.Drawing.Point(0, 0);
                if(this is IPairwiseSpec)
                {
                    NumericUpDown upDown = (NumericUpDown)CorrelVisual.Controls.Find("numericUpDown_CorrelValue", true).First();
                    upDown.Enabled = false;
                    Button drawButton = (Button)CorrelVisual.Controls.Find("btn_LaunchDrawCorrelation", true).First();
                    drawButton.Enabled = false;
                }
                CorrelVisual.ShowDialog();
                CorrelVisual.Focus();
            }

            public virtual void FormatSheet() { throw new Exception("Failed override"); }

            public override bool Validate() { throw new Exception("Failed override"); }

            public virtual void ConvertCorrelation( bool PreserveOffDiagonal=false) { throw new Exception("Failed override"); }

            protected void ConversionFormClicked(object sender, EventArgs e)      //This works.. but why? Isn't the object gone?
            {
                var conversionForm = new CorrelationConversionForm(this);
                conversionForm.Show();
                conversionForm.Focus();
            }

            protected void CollapseCorrelationClicked(object sender, EventArgs e)
            {
                CollapseToSheet();
            }

            protected void CancelChangesClicked(object sender, EventArgs e)
            {
                //Delete the sheet
                ThisAddIn.MyApp.DisplayAlerts = false;
                this.LinkToOrigin.LinkSource.Worksheet.Activate();
                this.xlSheet.Delete();
                ThisAddIn.MyApp.DisplayAlerts = true;
            }

            protected void VisualizeCorrelationClicked(object sender, EventArgs e)
            {
                Sheets.CorrelationSheet correlSheet = ConstructFromXlCorrelationSheet();
                if (correlSheet == null)
                    return;
                correlSheet.VisualizeCorrel();
            }
        }
    }
}
