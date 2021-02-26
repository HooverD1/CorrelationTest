using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace CorrelationTest
{
    namespace Data
    {
        public enum MatrixErrors
        {
            AboveUpperBound,
            BelowLowerBound,
            MisplacedValue,
            None
        }

        public class CorrelationMatrix
        {
            public Dictionary<string, int> FieldDict { get; set; }
            public object[,] Matrix { get; set; }
            private double[,] DoubleMatrix { get; set; }
            public int FieldCount { get; set; }
            public string[] Fields { get; set; }
            public string[] IDs { get; set; }
            public string Parent_ID { get; set; }
            public Sheets.CorrelationSheet ContainingSheet { get; set; }

            protected CorrelationMatrix() { }       //default

            protected CorrelationMatrix(double[,] phasingTriple)
            {
                //build a phasing correlation matrix from a provided triple
            }

            public CorrelationMatrix(Sheets.CorrelationSheet containingSheet, Excel.Range fieldsRange, Excel.Range matrixRange)       //from matrix
            {

                if (fieldsRange.Cells.Count != matrixRange.Columns.Count)
                    throw new Exception("Names do not match matrix.");
                this.ContainingSheet = containingSheet;
                this.FieldCount = fieldsRange.Cells.Count;
                Matrix = ExtensionMethods.ReIndexArray<object>(matrixRange.Value);
                object[,] fieldTemp = fieldsRange.Value;
                this.Fields = new string[fieldTemp.GetLength(1)];
                for(int i = 0; i < fieldTemp.GetLength(1); i++) { this.Fields[i] = fieldTemp[1, i + 1].ToString(); }
                FieldDict = GetFieldDict(fieldsRange, matrixRange);
            }

            public static CorrelationMatrix ConstructFromParentItem(IHasCorrelations ParentItem, SheetType correlType)
            {
                CorrelationMatrix matrix;
                //this should vary based on what type of CorrelationString!!
                switch (correlType)
                {
                    case SheetType.Correlation_CP:
                        matrix = new Data.CorrelationMatrix_Inputs();
                        matrix.IDs = (from ISub sub in ((IHasCostCorrelations)ParentItem).SubEstimates select sub.uID.ID).ToArray();
                        matrix.Fields = ParentItem.GetFields();
                        matrix.Matrix = ((IHasCostCorrelations)ParentItem).CostCorrelationString.GetMatrix(matrix.Fields);
                        matrix.FieldCount = matrix.Fields.Count();
                        matrix.FieldDict = matrix.GetFieldDict(matrix.IDs);
                        break;
                    case SheetType.Correlation_CM:
                        matrix = new Data.CorrelationMatrix_Inputs();
                        matrix.Fields = ParentItem.GetFields();
                        matrix.Matrix = ((IHasCostCorrelations)ParentItem).CostCorrelationString.GetMatrix(matrix.Fields);
                        matrix.FieldCount = matrix.Fields.Count();
                        matrix.IDs = (from ISub sub in ((IHasCostCorrelations)ParentItem).SubEstimates select sub.uID.ID).ToArray();
                        matrix.FieldDict = matrix.GetFieldDict(matrix.IDs);
                        break;
                    case SheetType.Correlation_PP:
                        matrix = new Data.CorrelationMatrix_Phasing();
                        matrix.Fields = (from Period period in ((IHasPhasingCorrelations)ParentItem).Periods select period.pID.ID).ToArray();
                        matrix.Matrix = ((IHasPhasingCorrelations)ParentItem).PhasingCorrelationString.GetMatrix(matrix.Fields);
                        matrix.FieldCount = matrix.Fields.Count();
                        matrix.IDs = (from Period sub in ((IHasPhasingCorrelations)ParentItem).Periods select sub.pID.ID).ToArray();
                        matrix.FieldDict = matrix.GetFieldDict(matrix.IDs);
                        break;
                    case SheetType.Correlation_PM:
                        matrix = new Data.CorrelationMatrix_Phasing();
                        matrix.Fields = (from Period period in ((IHasPhasingCorrelations)ParentItem).Periods select period.pID.ID).ToArray();
                        matrix.Matrix = ((IHasPhasingCorrelations)ParentItem).PhasingCorrelationString.GetMatrix(matrix.Fields);
                        matrix.FieldCount = matrix.Fields.Count();
                        matrix.IDs = (from Period sub in ((IHasPhasingCorrelations)ParentItem).Periods select sub.pID.ID).ToArray();
                        matrix.FieldDict = matrix.GetFieldDict(matrix.IDs);
                        break;
                    case SheetType.Correlation_DP:
                        matrix = new Data.CorrelationMatrix_Duration();
                        matrix.Fields = ParentItem.GetFields();
                        matrix.Matrix = ((IHasDurationCorrelations)ParentItem).DurationCorrelationString.GetMatrix(matrix.Fields);
                        matrix.FieldCount = matrix.Fields.Count();
                        matrix.IDs = (from ISub sub in ((IHasDurationCorrelations)ParentItem).SubEstimates select sub.uID.ID).ToArray();
                        matrix.FieldDict = matrix.GetFieldDict(matrix.IDs);
                        break;
                    case SheetType.Correlation_DM:
                        matrix = new Data.CorrelationMatrix_Duration();
                        matrix.Fields = ParentItem.GetFields();
                        matrix.Matrix = ((IHasDurationCorrelations)ParentItem).DurationCorrelationString.GetMatrix(matrix.Fields);
                        matrix.FieldCount = matrix.Fields.Count();
                        matrix.IDs = (from ISub sub in ((IHasDurationCorrelations)ParentItem).SubEstimates select sub.uID.ID).ToArray();
                        matrix.FieldDict = matrix.GetFieldDict(matrix.IDs);
                        break;
                    default:
                        throw new Exception("Invalid CorrelationString type.");
                }
                return matrix;
            }

            public static CorrelationMatrix ConstructFromCorrelationSheet(Sheets.CorrelationSheet CorrelSheet)
            {
                //Rebuilding the matrix object from the matrix on the xl sheet
                //Need to do this without building it into a string object
                string parent_ID = Convert.ToString(CorrelSheet.xlIDCell.Value);
                string[] sub_fields = (string[])CorrelSheet.Get_xlFields();
                string[] sub_IDs = CorrelSheet.GetIDs();

                //=======================
                //Pull the matrix from the xlSheet to get updates

                Excel.Range first_cell = CorrelSheet.xlMatrixCell.Offset[1, 0];
                Excel.Range last_cell = first_cell.End[Excel.XlDirection.xlToRight].End[Excel.XlDirection.xlDown];
                Excel.Range last_cell_fields = CorrelSheet.xlMatrixCell.End[Excel.XlDirection.xlToRight];
                object[,] matrix = CorrelSheet.xlSheet.Range[first_cell, last_cell].Value;
                //=======================

                CorrelationMatrix matrix_obj;

                SheetType containing_sheet_type = ExtensionMethods.GetSheetType(CorrelSheet.xlSheet);

                switch (containing_sheet_type)
                {
                    case SheetType.Correlation_CM:
                        matrix_obj = new CorrelationMatrix_Inputs();
                        matrix_obj.Parent_ID = parent_ID;
                        matrix_obj.IDs = sub_IDs;
                        matrix_obj.Fields = sub_fields;
                        matrix_obj.Matrix = matrix;
                        matrix_obj.FieldDict = matrix_obj.GetFieldDict(matrix_obj.IDs);
                        break;
                    case SheetType.Correlation_CP:
                        matrix_obj = new CorrelationMatrix_Inputs();
                        matrix_obj.Parent_ID = parent_ID;
                        matrix_obj.IDs = sub_IDs;
                        matrix_obj.Fields = sub_fields;
                        matrix_obj.Matrix = matrix;
                        matrix_obj.FieldDict = matrix_obj.GetFieldDict(matrix_obj.IDs);
                        break;
                    case SheetType.Correlation_PM:
                        matrix_obj = new CorrelationMatrix_Phasing();
                        matrix_obj.Parent_ID = parent_ID;
                        PeriodID[] pids = PeriodID.GeneratePeriodIDs(UniqueID.ConstructFromExisting(Convert.ToString(parent_ID)), sub_fields.Count());
                        matrix_obj.IDs = pids.Select(x => x.ID).ToArray();
                        matrix_obj.Fields = sub_fields;
                        matrix_obj.Matrix = matrix;
                        matrix_obj.FieldDict = matrix_obj.GetFieldDict(matrix_obj.IDs);
                        break;
                    case SheetType.Correlation_PP:
                        matrix_obj = new CorrelationMatrix_Phasing();
                        matrix_obj.Parent_ID = parent_ID;
                        PeriodID[] pids2 = PeriodID.GeneratePeriodIDs(UniqueID.ConstructFromExisting(Convert.ToString(parent_ID)), sub_fields.Count());
                        matrix_obj.IDs = pids2.Select(x => x.ID).ToArray();
                        matrix_obj.Fields = sub_fields;
                        matrix_obj.Matrix = matrix;
                        matrix_obj.FieldDict = matrix_obj.GetFieldDict(matrix_obj.IDs);
                        break;
                    case SheetType.Correlation_DM:
                        matrix_obj = new CorrelationMatrix_Duration();
                        matrix_obj.Parent_ID = parent_ID;
                        matrix_obj.IDs = sub_IDs;
                        matrix_obj.Fields = sub_fields;
                        matrix_obj.Matrix = matrix;
                        matrix_obj.FieldDict = matrix_obj.GetFieldDict(matrix_obj.IDs);
                        break;
                    case SheetType.Correlation_DP:
                        matrix_obj = new CorrelationMatrix_Duration();
                        matrix_obj.Parent_ID = parent_ID;
                        matrix_obj.IDs = sub_IDs;
                        matrix_obj.Fields = sub_fields;
                        matrix_obj.Matrix = matrix;
                        matrix_obj.FieldDict = matrix_obj.GetFieldDict(matrix_obj.IDs);
                        break;
                    default:
                        throw new Exception("Unknown correlation type");
                }
                matrix_obj.Matrix = ExtensionMethods.ReIndexArray(matrix);
                matrix_obj.FieldCount = matrix_obj.Fields.Count();

                return matrix_obj;
            }
                        
            private bool Even(int fieldCount)
            {
                if (fieldCount % 2 == 0)
                    return true;
                else
                    return false;

            }
            private int GetMidpoint(int fieldCount, bool isEven)
            {
                if (isEven)
                    return fieldCount / 2;
                else
                    return fieldCount / 2 + 1;
            }

            private Dictionary<string, int> GetFieldDict(Excel.Range fieldRange, Excel.Range matrixRange)       //get the field index by unique id
            {
                var fieldDict = new Dictionary<string, int>();      //<field name, index>
                //Excel.Range fieldEnd = fieldStart.Offset[0, matrixRange.Columns.Count];
                //Excel.Range fieldRange = matrixRange.Worksheet.Range[fieldStart, fieldEnd];
                object[,] fieldStrings = fieldRange.Value;  // field names
                //fieldStrings = fieldRange.Value2;
                Excel.Worksheet sourceSheet = ThisAddIn.MyApp.get_Range((object)this.ContainingSheet.xlLinkCell.Value).Worksheet;
                
                for (int i = 1; i <= this.FieldCount; i++)
                {
                    fieldDict.Add(fieldStrings[1,i].ToString(), i);      //is this being launched off correlation sheet? If so, have to follow the link
                }
                return fieldDict;
            }

            private Dictionary<string, int> GetFieldDict(string[] ids)
            {
                string[] id_strings = ids.Select(x=>Convert.ToString(x)).ToArray();
                FieldDict = new Dictionary<string, int>();
                for (int i = 0; i < ids.Count(); i++)
                {
                    if (!FieldDict.ContainsKey(id_strings[i]))
                        FieldDict.Add(id_strings[i], i);
                    else
                        throw new Exception("IDs are not unique");
                }
                return FieldDict;
            }

            public object[,] GetMatrix()
            {
                return Matrix;
            }

            public object[] GetFields()
            {
                return Fields;
            }

            private string ParseID(string id)
            {
                string[] id_pieces = id.Split('|');         //split lines
                if (id_pieces.Length == 2)
                    return id_pieces[1];                    //return the name portion of the ID
                else
                    return null;                            //if malformed, return null
            }

            public double AccessArray(string id1, string id2)
            {
                //Access values by unique id pairs
                if (id1.Equals(id2))
                    return 1;
                else
                    return Convert.ToDouble(Matrix[FieldDict[id1], FieldDict[id2]]);
            }

            public void SetCorrelation(string id1, string id2, double correlation)
            {
                Matrix[FieldDict[id1], FieldDict[id2]] = correlation;
            }
         
            public void PrintToSheet(Excel.Range xlRange)
            {
                xlRange.Resize[1, this.FieldDict.Count].Value = this.Fields;                                     //print fields
                object[,] transpose = new object[this.Fields.Length,1];
                for (int i = 0; i < this.Fields.Length; i++)
                    transpose[i, 0] = this.Fields[i];
                xlRange.Offset[1, -1].Resize[this.FieldCount, 1].Value = transpose;
                xlRange.Offset[1,0].Resize[Matrix.GetLength(0),Matrix.GetLength(1)].Value = this.Matrix;    //print matrix
            }
            public bool ValidateAgainstXlSheet(object[] xlSheetFields)
            {
                var localFields = this.Fields;
                if (localFields.Count() != xlSheetFields.Count())
                {
                    return false;
                }
                foreach (object field in localFields)
                {
                    if (!xlSheetFields.Contains<object>(field))
                        return false;
                }
                return true;
            }
            private double[,] GetDoubleMatrix(object[,] objectMatrix)
            {
                objectMatrix = ExtensionMethods.ReIndexArray<object>(objectMatrix);
                objectMatrix = ExtensionMethods.AddLowerTriangular(objectMatrix);
                if (this.DoubleMatrix == null)
                {
                    this.DoubleMatrix = new double[objectMatrix.GetLength(0), objectMatrix.GetLength(1)];
                    for (int row = 0; row < objectMatrix.GetLength(0); row++)
                    {
                        for (int col = 0; col < objectMatrix.GetLength(1); col++)
                        {
                            DoubleMatrix[row, col] = Convert.ToDouble(objectMatrix[row, col]);
                        }
                    }
                }
                return this.DoubleMatrix;
            }
            
            public bool CheckForPSD()
            {
                double[,] doubleMatrix = GetDoubleMatrix(this.Matrix);
                var eigens = new Accord.Math.Decompositions.EigenvalueDecomposition(doubleMatrix, false, true);
                if (eigens.RealEigenvalues.Min() < 0)
                    return false;
                else
                    return true;
            }

            public MatrixErrors[,] CheckMatrixForTransitivity()
            {
                
                this.DoubleMatrix = GetDoubleMatrix(this.Matrix);
                MatrixErrors[,] errorMatrix = new MatrixErrors[this.Matrix.GetLength(0), this.Matrix.GetLength(1)];

                for(int row = 0; row < errorMatrix.GetLength(0); row++)
                {
                    for(int col = 0; col < errorMatrix.GetLength(1); col++)
                    {
                        if (row == col)
                        {
                            if (DoubleMatrix[row, col] > 1)
                                errorMatrix[row, col] = MatrixErrors.AboveUpperBound;
                            else if (DoubleMatrix[row, col] < 1)
                                errorMatrix[row, col] = MatrixErrors.BelowLowerBound;
                            else
                                errorMatrix[row, col] = MatrixErrors.None;
                            continue;
                        }
                        if(col < row)
                        {
                            if (DoubleMatrix[row, col] == Convert.ToDouble(this.Matrix[col, row]))
                                errorMatrix[row, col] = MatrixErrors.None;
                            else
                                errorMatrix[row, col] = MatrixErrors.MisplacedValue;
                            continue;
                        }
                        double max_lower=-1;
                        double min_upper=1;
                        for(int via = 0; via < errorMatrix.GetLength(0); via++)
                        {
                            if (row == via || col == via)
                                continue;
                            double lowerBound = GetTransLowerBound(row, via, col);
                            double upperBound = GetTransUpperBound(row, via, col);
                            if (lowerBound > max_lower)
                                max_lower = lowerBound;
                            if (upperBound < min_upper)
                                min_upper = upperBound;
                            
                        }
                        if (min_upper < DoubleMatrix[row, col])
                        {
                            errorMatrix[row, col] = MatrixErrors.AboveUpperBound;
                        }
                        else if(max_lower > DoubleMatrix[row, col])
                        {
                            errorMatrix[row, col] = MatrixErrors.BelowLowerBound;
                        }
                        else if (max_lower <= DoubleMatrix[row, col] && min_upper >= DoubleMatrix[row, col])
                        {
                            errorMatrix[row, col] = MatrixErrors.None;
                        }
                        else
                        {
                            errorMatrix[row, col] = MatrixErrors.None;
                        }
                    }
                }
                return errorMatrix;
            }
            private double GetTransLowerBound(int x, int y, int z)
            {
                double Pxy = this.DoubleMatrix[x, y];
                double Pyz = this.DoubleMatrix[y, z];
                return Pxy * Pyz - Math.Sqrt((1 - Pxy * Pxy) * (1 - Pyz * Pyz));
            }
            private double GetTransUpperBound(int x, int y, int z)
            {
                double Pxy = this.DoubleMatrix[x, y];
                double Pyz = this.DoubleMatrix[y, z];
                return Pxy * Pyz + Math.Sqrt((1 - Pxy * Pxy) * (1 - Pyz * Pyz));
            }

            public bool Equals(object[,] cm)
            {
                if (cm.GetLength(0) != this.Matrix.GetLength(0) || cm.GetLength(1) != this.Matrix.GetLength(1))
                    return false;
                for(int row = 0; row < this.Matrix.GetLength(0); row++)
                {
                    for(int col = row; col < this.Matrix.GetLength(1); col++)
                    {
                        if (!Double.TryParse(this.Matrix[row, col].ToString(), out double internalVal))
                            throw new Exception("Invalid matrix value");
                        if (!Double.TryParse(cm[row, col].ToString(), out double externalVal))
                            throw new Exception("Invalid matrix value");
                        if (internalVal != externalVal)
                            return false;
                    }
                }
                return true;
            }

            public bool ValidateAgainstTriple(Triple pt)
            {
                object[,] tripleMatrix = pt.GetCorrelationMatrix(this.Fields);
                return this.Equals(tripleMatrix);
            }
        }   //class
    }//Data
}//namespace
