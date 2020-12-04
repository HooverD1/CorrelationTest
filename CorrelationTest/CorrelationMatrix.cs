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
            public Dictionary<UniqueID, int> FieldDict { get; set; }
            private object[,] PartialArray { get; set; }
            private object[,] SecondaryArray { get; set; }
            private object[,] Matrix { get; set; }
            private double[,] DoubleMatrix { get; set; }
            private int Midpoint { get; set; }
            private bool IsEven { get; set; }
            public int FieldCount { get; set; }
            public object[] Fields { get; set; }
            public Tuple<int, int> MatrixCoords { get; }

            public CorrelationMatrix(Excel.Range correlMatrix)       //from matrix
            {
                this.FieldCount = correlMatrix.Columns.Count;
                this.IsEven = Even(this.FieldCount);
                this.Midpoint = GetMidpoint(this.FieldCount, this.IsEven);
                Matrix = GetMainRange(correlMatrix);
                SecondaryArray = GetSecondaryRange(correlMatrix);
                FieldDict = GetFieldDict(correlMatrix);
            }

            public CorrelationMatrix(Data.CorrelationString_Inputs correlStringObj)
            {
                //expand from string
                this.Fields = correlStringObj.GetFields();
                this.Matrix = correlStringObj.GetMatrix();      //creates a correlation matrix & loops
                this.FieldCount = this.Fields.Count();
                this.IsEven = Even(this.FieldCount);
                this.Midpoint = GetMidpoint(this.FieldCount, this.IsEven);
                this.FieldDict = GetFieldDict(correlStringObj.GetIDs());
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

            private object[,] GetMainRange(string[] fields, double[,] matrix)
            {
                object[,] mainRange = new object[this.Midpoint, this.Midpoint];
                for (int row = 0; row < this.Midpoint; row++)
                {
                    for (int col = 0; col < this.Midpoint; col++)
                    {
                        int col2 = this.Midpoint - col;
                        mainRange[row, col] = matrix[row, col2];
                    }
                }
                return mainRange;
            }

            private object[,] GetMainRange(Excel.Range myRange)
            {
                Excel.Range mainRange;
                Excel.Range mainFirstCell;
                Excel.Range mainLastCell;
                if (this.IsEven == true)
                {
                    mainFirstCell = myRange.Cells[1, this.Midpoint];
                    mainLastCell = myRange.Cells[this.Midpoint, this.FieldCount];
                }
                else
                {
                    mainFirstCell = myRange.Offset[0, this.Midpoint + 1];
                    mainLastCell = myRange.Offset[this.Midpoint - 1, this.FieldCount];
                }
                mainRange = myRange.Range[mainFirstCell, mainLastCell];
                return mainRange.Value2;
            }

            private object[,] GetSecondaryRange(Excel.Range myRange)
            {
                Excel.Range secondRange;
                Excel.Range secondFirstCell;
                Excel.Range secondLastCell;
                if (this.IsEven == true)
                {
                    secondFirstCell = myRange.Cells[1, 1];
                    secondLastCell = myRange.Cells[this.Midpoint, this.Midpoint];
                }
                else
                {
                    secondFirstCell = myRange.Offset[1, 1];
                    secondLastCell = myRange.Offset[this.Midpoint, this.Midpoint];
                }
                secondRange = myRange.Range[secondFirstCell, secondLastCell];
                object[,] topLeft = secondRange.Value2;
                if (this.IsEven == true)
                {
                    secondFirstCell = myRange.Cells[this.Midpoint + 1, this.Midpoint + 1];
                    secondLastCell = myRange.Cells[this.FieldCount, this.FieldCount];
                }
                else
                {
                    secondFirstCell = myRange.Offset[this.Midpoint + 1, this.Midpoint + 1];
                    secondLastCell = myRange.Offset[this.FieldCount, this.FieldCount];
                }
                secondRange = myRange.Range[secondFirstCell, secondLastCell];
                object[,] bottomRight = secondRange.Value2;
                for (int row = this.Midpoint + 1; row < this.FieldCount; row++)
                {
                    for (int col = row + 1; col < this.FieldCount; col++)
                    {
                        var coords = this.TransformField(row, col);
                        topLeft[coords.Item2, coords.Item3] = bottomRight[row - this.Midpoint, col - this.Midpoint];
                    }
                }
                return topLeft;
            }

            private Dictionary<UniqueID, int> GetFieldDict(Excel.Range myRange)       //get the field index by unique id
            {
                FieldDict = new Dictionary<UniqueID, int>();
                Excel.Range fieldStart = myRange.Offset[-1, 0];
                Excel.Range fieldEnd = myRange.Offset[myRange.Columns.Count, 0];
                Excel.Range fieldRange = myRange.Worksheet.Range[fieldStart, fieldEnd];
                object[,] fieldStrings = new object[1, this.FieldCount];
                fieldStrings = fieldRange.Value2;
                for (int i = 1; i <= this.FieldCount; i++)
                {
                    FieldDict.Add(new UniqueID(myRange.Worksheet.Name, fieldStrings[1, i].ToString()), i);      //is this being launched off correlation sheet? If so, have to follow the link
                }
                return FieldDict;
            }

            private Dictionary<UniqueID, int> GetFieldDict(UniqueID[] ids)
            {
                FieldDict = new Dictionary<UniqueID, int>();
                for (int i = 0; i < ids.Count(); i++)
                {
                    if (!FieldDict.ContainsKey(ids[i]))
                        FieldDict.Add(ids[i], i);
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
                return FieldDict.Keys.Select(x => x.Name).ToArray<object>();
            }

            public UniqueID[] GetIDs()
            {
                return FieldDict.Keys.ToArray();
            }

            private string ParseID(string id)
            {
                string[] id_pieces = id.Split('|');         //split lines
                if (id_pieces.Length == 2)
                    return id_pieces[1];                    //return the name portion of the ID
                else
                    return null;                            //if malformed, return null
            }

            public double AccessArray(UniqueID id1, UniqueID id2)
            {
                //Access values by unique id pairs
                if (id1.Equals(id2))
                    return 1;
                else
                    return Convert.ToDouble(Matrix[FieldDict[id1], FieldDict[id2]]);
            }

            public void SetCorrelation(UniqueID id1, UniqueID id2, double correlation)
            {
                Matrix[FieldDict[id1], FieldDict[id2]] = correlation;
            }

            private enum ArrayType
            {
                Main,
                Secondary
            }
            private Tuple<ArrayType, int, int> TransformField(int rowIndex, int colIndex)
            {
                //Take a field name, check it's index with the dictionary, and transform it if need be

                if (this.IsEven == true)
                {
                    if (rowIndex < this.Midpoint)
                    {
                        if (colIndex > this.Midpoint)
                        {
                            //top right quadrant
                            return new Tuple<ArrayType, int, int>(ArrayType.Main, rowIndex, colIndex);
                        }
                        else
                        {
                            //top left quadrant
                            return new Tuple<ArrayType, int, int>(ArrayType.Secondary, rowIndex, colIndex);
                        }
                    }
                    else
                    {
                        if (colIndex > this.Midpoint)
                        {
                            //bottom right quadrant
                            int newRowIndex = this.FieldCount - rowIndex;
                            int newColIndex = this.FieldCount - colIndex;
                            return new Tuple<ArrayType, int, int>(ArrayType.Secondary, newRowIndex, newColIndex);
                        }
                        else
                        {
                            //bottom left quadrant
                            //Convert to top right
                            return new Tuple<ArrayType, int, int>(ArrayType.Main, this.FieldCount - rowIndex, this.FieldCount - colIndex);
                        }
                    }
                }
                else
                {
                    throw new NotImplementedException();
                }
            }
            public void PrintToSheet(Excel.Range xlRange)
            {
                xlRange.Resize[1, this.FieldCount].Value = this.Fields;                                     //print fields
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
                            if (this.Matrix[row,col] == null)
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
        }   //class
    }//Data
}//namespace
