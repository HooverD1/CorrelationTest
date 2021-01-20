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

            public Data.CorrelationString CorrelString { get; set; }
            protected Data.CorrelationMatrix CorrelMatrix { get; set; }
            public Excel.Range xlMatrixCell { get; set; }
            public Data.Link LinkToOrigin { get; set; }
            public Excel.Range xlLinkCell { get; set; }
            public Excel.Range xlIDCell { get; set; }
            public Excel.Range xlCorrelStringCell { get; set; }
            public Excel.Range xlDistCell { get; set; }
            public Data.CorrelSheetSpecs Specs { get; set; }

            //public CorrelationSheet(Data.CorrelationString_IM correlString, Excel.Range launchedFrom) : this(correlString, launchedFrom, new Data.CorrelSheetSpecs()) { }       //default locations


            protected virtual Excel.Worksheet CreateXLCorrelSheet(string postfix) { throw new Exception("Failed override"); }
            protected virtual Excel.Worksheet GetXlSheet(bool CreateNew = true) { throw new Exception("Failed override"); }
            protected virtual Excel.Worksheet GetXlSheet(SheetType sheetType, bool CreateNew = true) { throw new Exception("Failed override"); }
            public virtual void UpdateCorrelationString(string[] ids) { throw new Exception("Failed override"); }

            public virtual string[] GetIDs()
            {
                string[] lines = Data.CorrelationString.DelimitString(Convert.ToString(this.xlCorrelStringCell.Value));
                string[] header = lines[0].Split(',');
                string[] ids = new string[header.Length - 3];
                for (int i = 3; i < header.Length; i++)
                    ids[i - 3] = header[i];
                return ids;
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
            protected void PrintMatrixEndCoords(Excel.Worksheet xlCorrelSheet)
            {
                this.Specs.MatrixCoords_End = GetMatrixEndCoords(xlMatrixCell, this.CorrelMatrix.FieldCount);
                xlCorrelSheet.Cells[Sheets.CorrelationSheet.param_RowMatrix2.Item1, Sheets.CorrelationSheet.param_RowMatrix2.Item2].Value = this.Specs.MatrixCoords_End.Item1;  //prints the row coord
                xlCorrelSheet.Cells[Sheets.CorrelationSheet.param_ColMatrix2.Item1, Sheets.CorrelationSheet.param_ColMatrix2.Item2].Value = this.Specs.MatrixCoords_End.Item2;  //prints the col coord
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
                for(int subIndex = 0; subIndex < tempEst.SubEstimates.Count(); subIndex++)      //Print the Distribution strings
                {
                    this.xlDistCell.Offset[subIndex, 0].Value = GetDistributionString(tempEst, subIndex);
                }                
            }

            protected string GetDistributionString(Estimate_Item est, int subIndex)
            {
                StringBuilder sb = new StringBuilder();
                sb.Append($"{est.SubEstimates[subIndex].ItemDistribution.Name}");
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

            public static CorrelationSheet Construct()      //Overload for Building object from EXISTING correlation xlsheet -- used for correl collapse
            {
                Excel.Worksheet xlCorrelSheet;
                List<Excel.Worksheet> xlCorrelSheets = new List<Excel.Worksheet>();
                foreach(Excel.Worksheet sht in ThisAddIn.MyApp.Worksheets)
                {
                    SheetType sht_type = ExtensionMethods.GetSheetType(sht);
                    if(sht_type == SheetType.Correlation_DM ||
                        sht_type == SheetType.Correlation_DT ||
                        sht_type == SheetType.Correlation_IT ||
                        sht_type == SheetType.Correlation_IM ||
                        sht_type == SheetType.Correlation_PM || 
                        sht_type == SheetType.Correlation_PT)
                    {
                        xlCorrelSheets.Add(sht);
                    }                    
                }
                if (xlCorrelSheets.Count() == 1)
                    xlCorrelSheet = xlCorrelSheets.First();
                else if (xlCorrelSheets.Count() > 1)
                    throw new Exception("Multiple correlation sheets");
                else if (!xlCorrelSheets.Any())
                    throw new Exception("No correlation sheets");
                else
                    throw new Exception("Unknown error finding correlation sheet");

                //Need to build the components from the xlSheet here instead of in the constructor, then build the sheet using .Construct()
                //CorrelString, Excel.Range source, new CorrelSheetSpecs()
                
                SheetType sheet_type = ExtensionMethods.GetSheetType(xlCorrelSheet);
                Data.CorrelSheetSpecs csSpecs = new Data.CorrelSheetSpecs(sheet_type);
                //Data.CorrelationString cs = Data.CorrelationString.Construct(xlCorrelSheet.Cells[csSpecs.StringCoords.Item1, csSpecs.StringCoords.Item2].value);
                Excel.Range source = ThisAddIn.MyApp.get_Range((object)xlCorrelSheet.Cells[csSpecs.LinkCoords.Item1, csSpecs.LinkCoords.Item2].value);
                //Make the sheetid cell in 1,1 list the type of correlation
                CorrelationSheet newSheet;
                switch (sheet_type)
                {
                    case SheetType.Correlation_DM:
                        throw new NotImplementedException();
                    case SheetType.Correlation_IM:
                        newSheet = new CorrelationSheet_Inputs(csSpecs);
                        break;
                    case SheetType.Correlation_PM:
                        newSheet = new CorrelationSheet_Phasing(csSpecs);
                        break;
                    case SheetType.Correlation_PT:
                        newSheet = new CorrelationSheet_Phasing(csSpecs);
                        break;
                    default:
                        throw new Exception("Not a valid Correlation Sheet type");
                }

                newSheet.xlSheet = xlCorrelSheet;
                newSheet.xlLinkCell = newSheet.xlSheet.Cells[csSpecs.LinkCoords.Item1, csSpecs.LinkCoords.Item2];
                newSheet.xlCorrelStringCell = newSheet.xlSheet.Cells[csSpecs.StringCoords.Item1, csSpecs.StringCoords.Item2];
                newSheet.xlIDCell = newSheet.xlSheet.Cells[csSpecs.IdCoords.Item1, csSpecs.IdCoords.Item2];
                newSheet.xlDistCell = newSheet.xlSheet.Cells[csSpecs.DistributionCoords.Item1, csSpecs.DistributionCoords.Item2];
                //need to be able to parse a link into sheetname and address to reconstruct the linkSource sheet
                newSheet.LinkToOrigin = new Data.Link(newSheet.xlLinkCell.Value);     //bootstrap the Link from the address
                newSheet.xlMatrixCell = newSheet.xlSheet.Cells[csSpecs.MatrixCoords.Item1, csSpecs.MatrixCoords.Item2];
                //Excel.Range matrix_end = newSheet.xlMatrixCell.End[Excel.XlDirection.xlToRight].End[Excel.XlDirection.xlDown];
                //Excel.Range matrixRange = newSheet.xlSheet.Range[newSheet.xlMatrixCell.Offset[1,0], matrix_end];
                //newSheet.CorrelMatrix = new Data.CorrelationMatrix(newSheet, newSheet.xlMatrixCell.Resize[1,matrixRange.Columns.Count], matrixRange.Offset[1,0].Resize[matrixRange.Columns.Count, matrixRange.Columns.Count]);
                //newSheet.UpdateCorrelationString(newSheet.GetIDs());     //Updates the string off the matrix

                return newSheet;
            }

            private Data.CorrelStringType GetCorrelType(string correlStringValue)
            {
                switch (correlStringValue)
                {
                    case "IT":
                        return Data.CorrelStringType.InputsTriple;
                    case "IM":
                        return Data.CorrelStringType.InputsMatrix;
                    case "PM":
                        return Data.CorrelStringType.PhasingMatrix;
                    case "PT":
                        return Data.CorrelStringType.PhasingTriple;
                    case "DT":
                        return Data.CorrelStringType.DurationTriple;
                    case "DM":
                        return Data.CorrelStringType.DurationMatrix;
                    default:
                        throw new Exception("Malformed ID");
                }
            }

            //public virtual void UpdateCorrelationString() { throw new Exception("Failed override"); }

            public static void CollapseToSheet()    //grab the xlSheet matrix, build the correlString from it, place it at the origin, delete the xlSheet
            {
                CorrelationSheet correlSheet = Construct();
                if (correlSheet == null)
                    return;

                //Validate matrix checks
                //Validate link source ID
                //Validate that the linkSource still has an ID match. If so, .PrintToSheet ... Otherwise, search for the ID and throw a warning ... if no ID can be found, throw an error and don't delete the sheet
                CostSheet originSheet = CostSheet.Construct(correlSheet.LinkToOrigin.LinkSource.Worksheet);
                object id_followLink = correlSheet.LinkToOrigin.LinkSource.EntireRow.Cells[1, originSheet.Specs.ID_Offset].value;
                object id_correlSheet = correlSheet.xlIDCell.Value;
                
                if (id_followLink.ToString() == id_correlSheet.ToString())
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

                correlSheet.LinkToOrigin.LinkSource.Worksheet.Activate();
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

            public static CorrelationSheet Construct(Data.CorrelationString correlString, Excel.Range source, Data.CorrelSheetSpecs specs)       //CorrelationSheet dynamic creator
            {
                if(correlString is Data.CorrelationString_IM)
                {
                    return new CorrelationSheet_Inputs((Data.CorrelationString_IM)correlString, source, specs);
                }
                else if(correlString is Data.CorrelationString_IT)
                {
                    return new CorrelationSheet_Inputs((Data.CorrelationString_IT)correlString, source, specs);
                }
                else if(correlString is Data.CorrelationString_PM)
                {
                    return new CorrelationSheet_Phasing((Data.CorrelationString_PM)correlString, source, specs);
                }
                else if (correlString is Data.CorrelationString_PT)
                {
                    return new CorrelationSheet_Phasing((Data.CorrelationString_PT)correlString, source, specs);
                }
                else
                {
                    throw new Exception("Unknown correlation string type.");
                }
            }

        }
    }
}
