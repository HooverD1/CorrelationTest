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

            //public Excel.Worksheet xlSheet { get; set; }
            private Data.CorrelationString CorrelString { get; set; }
            private Data.CorrelationMatrix CorrelMatrix { get; set; }
            private Excel.Range xlMatrixCell { get; set; }
            private Data.Link LinkToOrigin { get; set; }
            private Excel.Range xlLinkCell { get; set; }
            private Excel.Range xlIDCell { get; set; }
            private Excel.Range xlDistCell { get; set; }
            private Data.CorrelSheetSpecs Specs { get; set; }

            public CorrelationSheet(Data.CorrelationString correlString, Excel.Range launchedFrom) : this(correlString, launchedFrom, new Data.CorrelSheetSpecs()) { }       //default locations
            public CorrelationSheet(Data.CorrelationString correlString, Excel.Range launchedFrom, Data.CorrelSheetSpecs specs)        //bring in the coordinates and set up the ranges once they exist
            {
                this.CorrelString = correlString;
                this.Specs = specs;
                var xlCorrelSheets = from Excel.Worksheet sheet in ThisAddIn.MyApp.Worksheets
                                    where sheet.Cells[1, 1].Value == "$Correlation"
                                    select sheet;
                if (xlCorrelSheets.Any())
                    xlSheet = xlCorrelSheets.First();
                else
                    xlSheet = CreateXLCorrelSheet();
                CorrelMatrix = new Data.CorrelationMatrix(correlString);
                this.LinkToOrigin = new Data.Link(launchedFrom);
                this.xlLinkCell = xlSheet.Cells[specs.LinkCoords.Item1, specs.LinkCoords.Item2];
                this.xlIDCell = xlSheet.Cells[specs.IdCoords.Item1, specs.IdCoords.Item2];
                this.xlDistCell = xlSheet.Cells[specs.DistributionCoords.Item1, specs.IdCoords.Item2];
                this.xlMatrixCell = xlSheet.Cells[specs.MatrixCoords.Item1, specs.MatrixCoords.Item2];
                this.Specs.PrintMatrixCoords(xlSheet);                                          //Print the matrix start coords
                this.PrintMatrixEndCoords(xlSheet);                                             //Print the matrix end coords
                this.Specs.PrintLinkCoords(xlSheet);                                            //Print the link coords
                this.Specs.PrintIdCoords(xlSheet);                                              //Print the ID coords
                this.Specs.PrintDistCoords(xlSheet);                                            //Print the Distribution coords
            }
            private CorrelationSheet(Excel.Worksheet xlSheet)  //constructor for collapsing existing sheet back to a string and for working with the existing xlSheet
            {
                this.xlSheet = xlSheet;                
                int matrixRow = Convert.ToInt32(xlSheet.Cells[CorrelationSheet.param_RowMatrix.Item1, CorrelationSheet.param_RowMatrix.Item2].Value);
                int matrixCol = Convert.ToInt32(xlSheet.Cells[CorrelationSheet.param_ColMatrix.Item1, CorrelationSheet.param_ColMatrix.Item2].Value);
                int matrixRow_end = Convert.ToInt32(xlSheet.Cells[CorrelationSheet.param_RowMatrix2.Item1, CorrelationSheet.param_RowMatrix2.Item2].Value);
                int matrixCol_end = Convert.ToInt32(xlSheet.Cells[CorrelationSheet.param_ColMatrix2.Item1, CorrelationSheet.param_ColMatrix2.Item2].Value);
                int linkRow = Convert.ToInt32(xlSheet.Cells[CorrelationSheet.param_RowLink.Item1, CorrelationSheet.param_RowLink.Item2].Value);
                int linkCol = Convert.ToInt32(xlSheet.Cells[CorrelationSheet.param_ColLink.Item1, CorrelationSheet.param_ColLink.Item2].Value);
                int idRow = Convert.ToInt32(xlSheet.Cells[CorrelationSheet.param_RowID.Item1, CorrelationSheet.param_RowID.Item2].Value);
                int idCol = Convert.ToInt32(xlSheet.Cells[CorrelationSheet.param_ColID.Item1, CorrelationSheet.param_ColID.Item2].Value);
                int distRow = Convert.ToInt32(xlSheet.Cells[CorrelationSheet.param_RowDist.Item1, CorrelationSheet.param_RowDist.Item2].Value);
                int distCol = Convert.ToInt32(xlSheet.Cells[CorrelationSheet.param_ColDist.Item1, CorrelationSheet.param_ColDist.Item2].Value);
                this.Specs = new Data.CorrelSheetSpecs(matrixRow, matrixCol, linkRow, linkCol);
                this.xlLinkCell = xlSheet.Cells[linkRow, linkCol];
                this.xlIDCell = xlSheet.Cells[idRow, idCol];
                this.xlDistCell = xlSheet.Cells[distRow, distCol];
                //need to be able to parse a link into sheetname and address to reconstruct the linkSource sheet
                this.LinkToOrigin = new Data.Link(xlLinkCell.Value);     //bootstrap the Link from the address
                this.xlMatrixCell = xlSheet.Cells[matrixRow, matrixCol];
                Excel.Range matrix_end = xlSheet.Cells[matrixRow_end, matrixCol_end];
                Excel.Range matrixRange = xlSheet.Range[xlMatrixCell, matrix_end];
                this.CorrelString = new Data.CorrelationString(matrixRange);
                this.CorrelMatrix = new Data.CorrelationMatrix(this.CorrelString);
            }

            private Excel.Worksheet CreateXLCorrelSheet()
            {
                Excel.Worksheet xlCorrelSheet = ThisAddIn.MyApp.Worksheets.Add(After: ThisAddIn.MyApp.ActiveWorkbook.Sheets[ThisAddIn.MyApp.ActiveWorkbook.Sheets.Count]);
                xlCorrelSheet.Name = "Correlation";
                xlCorrelSheet.Cells[1, 1] = "$Correlation";
                xlCorrelSheet.Rows[1].Hidden = true;
                return xlCorrelSheet;
            }

           
            public Data.CorrelationString CollapseToString(object[,] correlArray)
            {
                throw new NotImplementedException();
            }
            public double[,] GetCorrelArray(Excel.Range xlRange)
            {
                throw new NotImplementedException();
            }

            public double[,] SetCorrelArray(Data.CorrelationMatrix correlMatrix)
            {
                throw new NotImplementedException();
            }
            public object[] Get_xlFields()      //error
            {
                int colNum = Convert.ToInt32(xlSheet.Cells[1, 9].Value);
                Excel.Range fieldRange = this.xlMatrixCell.Resize[1, Convert.ToInt32(colNum)];
                return fieldRange.Value;
            }
            public override bool Validate()
            {
                bool validateMatrix_to_String = this.CorrelString.ValidateAgainstMatrix(this.CorrelMatrix.Fields);
                //need to get fields from xlSheet fresh, not the object, to validate
                bool validateMatrix_to_xlSheet = this.CorrelMatrix.ValidateAgainstXlSheet(this.Get_xlFields());  
                return validateMatrix_to_String && validateMatrix_to_xlSheet;
            }
            private Tuple<int, int> GetMatrixEndCoords(Excel.Range xlMatrixCell, int fieldCount)
            {
                return new Tuple<int, int>(xlMatrixCell.Row + fieldCount, xlMatrixCell.Column + fieldCount-1);
            }
            private void PrintMatrixEndCoords(Excel.Worksheet xlCorrelSheet)
            {
                Tuple<int, int> MatrixEndCoords = GetMatrixEndCoords(xlMatrixCell, this.CorrelMatrix.FieldCount);
                xlCorrelSheet.Cells[Sheets.CorrelationSheet.param_RowMatrix2.Item1, Sheets.CorrelationSheet.param_RowMatrix2.Item2].Value = MatrixEndCoords.Item1;  //prints the row coord
                xlCorrelSheet.Cells[Sheets.CorrelationSheet.param_ColMatrix2.Item1, Sheets.CorrelationSheet.param_ColMatrix2.Item2].Value = MatrixEndCoords.Item2;  //prints the col coord
            }
            public override void PrintToSheet()  //expanding from string
            {
                Estimate tempEst = new Estimate(this.LinkToOrigin.LinkSource.EntireRow);        //Load only this parent estimate
                tempEst.LoadSubEstimates(this.LinkToOrigin.LinkSource.EntireRow);                //Load the sub-estimates for this estimate
                this.CorrelMatrix.PrintToSheet(xlMatrixCell);                                   //Print the matrix
                this.LinkToOrigin.PrintToSheet(xlLinkCell);                                     //Print the link
                this.xlIDCell.Value = tempEst.ID;                                               //Print the ID
                for(int subIndex = 0; subIndex < tempEst.SubEstimates.Count(); subIndex++)      //Print the Distribution strings
                {
                    this.xlDistCell.Offset[subIndex, 0].Value = GetDistributionString(tempEst, subIndex);
                }                
            }

            private string GetDistributionString(Estimate est, int subIndex)
            {
                StringBuilder sb = new StringBuilder();
                sb.Append($"{est.SubEstimates[subIndex].EstimateDistribution.Name}");
                for(int i = 1; i < est.SubEstimates[subIndex].DistributionParameters.Count(); i++)
                {
                    string param = $"Param{i}";
                    if (est.SubEstimates[subIndex].DistributionParameters[param] != null)
                        sb.Append($",{est.SubEstimates[subIndex].DistributionParameters[param]}");
                }
                return sb.ToString();
            }

            public bool PaintMatrixErrors(Data.MatrixErrors[,] matrixErrors)        //return false if no matrix errors found
            {
                IEnumerable<Data.MatrixErrors> findErrors = from Data.MatrixErrors me in matrixErrors
                                                            where me != Data.MatrixErrors.None
                                                            select me;
                if (findErrors.Any())
                {
                    for (int row = 0; row < matrixErrors.GetLength(0); row++)
                    {
                        for (int col = 0; col < matrixErrors.GetLength(1); col++)
                        {
                            if (matrixErrors[row, col] == Data.MatrixErrors.None)
                            {
                                this.xlMatrixCell.Offset[row + 1, col].ClearComments();
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

            public static CorrelationSheet BuildFromExisting()
            {
                //default to grabbing the correlation sheet matrix range, convert to string (inside private sheet constructor), and dropping it in the linked location
                Excel.Worksheet xlSheet;
                var xlCorrelSheets = from Excel.Worksheet sheet in ThisAddIn.MyApp.Worksheets
                                     where sheet.Cells[1, 1].Value == "$Correlation"
                                     select sheet;
                if (xlCorrelSheets.Any())
                    xlSheet = xlCorrelSheets.First();
                else
                    return null;
                return new CorrelationSheet(xlSheet);  //bootstrap off the xlSheet
            }

            public static void CollapseToSheet()    //grab the xlSheet matrix, build the correlString from it, place it at the origin, delete the xlSheet
            {
                CorrelationSheet correlSheet = BuildFromExisting();
                if (correlSheet == null)
                    return;
                //validate that the linkSource still has an ID match. If so, .PrintToSheet ... Otherwise, search for the ID and throw a warning ... if no ID can be found, throw an error and don't delete the sheet
                if (new Estimate(correlSheet.LinkToOrigin.LinkSource.EntireRow).GetID() == correlSheet.xlIDCell.Value)
                {
                    correlSheet.CorrelString.PrintToSheet(correlSheet.LinkToOrigin.LinkSource);
                    if (!correlSheet.CorrelMatrix.CheckForPSD())
                        MessageBox.Show("Not PSD");
                    bool errors = correlSheet.PaintMatrixErrors(correlSheet.CorrelMatrix.CheckMatrixForTransitivity());
                    if (!errors)
                    {
                        ThisAddIn.MyApp.DisplayAlerts = false; 
                        correlSheet.xlSheet.Delete();
                        ThisAddIn.MyApp.DisplayAlerts = true;
                    }                    
                }                    
                else
                    MessageBox.Show("ID not found");
            }

            private object[,] GetDistributionParamStrings()
            {
                int lastRow = xlDistCell.End[Excel.XlDirection.xlDown].Row;
                Excel.Range distRange = xlDistCell.Resize[lastRow - xlDistCell.Row+1, 1];
                return distRange.Value;
            }

            public void VisualizeCorrel()
            {
                //Select a correlation
                int selectionRow = ThisAddIn.MyApp.Selection.Row;       //dependency..
                int selectionCol = ThisAddIn.MyApp.Selection.Column;       //dependency..
                //Make sure the row and column are within the matrix range
                if (selectionRow < this.Specs.MatrixCoords.Item1 + 1 || selectionRow > this.CorrelMatrix.FieldCount + this.Specs.MatrixCoords.Item1)
                    return;
                if (selectionCol < this.Specs.MatrixCoords.Item2 || selectionCol > this.CorrelMatrix.FieldCount + this.Specs.MatrixCoords.Item2-1)
                    return;
                //Find the distributions to use
                object[,] distParams = GetDistributionParamStrings();
                int distRow = selectionRow - xlMatrixCell.Row;
                int distCol = selectionCol - xlMatrixCell.Column+1;
                //Create distribution objects
                Distribution d1 = new Distribution(distParams[distRow, 1].ToString());
                Distribution d2 = new Distribution(distParams[distCol, 1].ToString());
                //Create the form
                CorrelationForm CorrelVisual = new CorrelationForm(d1, d2);
                CorrelVisual.Show();
            }
        }
    }
}
