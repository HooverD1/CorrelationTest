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
            public object[,] Matrix { get; set; }
            private double[,] DoubleMatrix { get; set; }
            public int FieldCount { get; set; }
            public object[] Fields { get; set; }
            public Sheets.CorrelationSheet ContainingSheet { get; set; }

            private CorrelationMatrix(double[,] phasingTriple)
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
                this.Fields = new object[fieldTemp.GetLength(1)];
                for(int i = 0; i < fieldTemp.GetLength(1); i++) { this.Fields[i] = fieldTemp[1, i + 1]; }
                FieldDict = GetFieldDict(fieldsRange, matrixRange);
            }

            public CorrelationMatrix(Data.CorrelationString_Inputs correlStringObj)
            {
                //expand from string
                this.Fields = correlStringObj.GetFields();
                this.Matrix = correlStringObj.GetMatrix();      //creates a correlation matrix & loops
                this.FieldCount = this.Fields.Count();
                this.FieldDict = GetFieldDict(correlStringObj.GetIDs());
            }

            public CorrelationMatrix(Data.CorrelationString_Periods correlStringObj)
            {
                //expand from string
                this.Fields = correlStringObj.GetFields();
                this.Matrix = correlStringObj.GetMatrix();      //creates a correlation matrix & loops
                this.FieldCount = this.Fields.Count();
                this.FieldDict = GetFieldDict(correlStringObj.GetIDs());
            }

            public CorrelationMatrix(Data.CorrelationString_Triple correlStringObj)
            {
                //expand from string
                this.Fields = correlStringObj.GetFields();
                this.Matrix = correlStringObj.GetMatrix();      //creates a correlation matrix & loops
                this.FieldCount = this.Fields.Count();
                PeriodID[] pids = PeriodID.GeneratePeriodIDs(correlStringObj.GetIDs().First(), FieldCount);
                this.FieldDict = GetFieldDict(pids);
            }

            public CorrelationMatrix(UniqueID parent_uid, object[,] matrix)     //used for creating phasing correlation matrices
            {       //THIS NEEDS TO SET UP FIELD DICT
                this.Matrix = matrix;
                this.FieldCount = Matrix.GetLength(0);
                //validate parent_uid and matrix
                PeriodID[] pids = PeriodID.GeneratePeriodIDs(parent_uid, FieldCount);
                this.Fields = null; //No names in IDs anymore..
                this.FieldDict = GetFieldDict(pids);

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

            private Dictionary<UniqueID, int> GetFieldDict(Excel.Range fieldRange, Excel.Range matrixRange)       //get the field index by unique id
            {
                var fieldDict = new Dictionary<UniqueID, int>();      //<field name, index>
                //Excel.Range fieldEnd = fieldStart.Offset[0, matrixRange.Columns.Count];
                //Excel.Range fieldRange = matrixRange.Worksheet.Range[fieldStart, fieldEnd];
                object[,] fieldStrings = fieldRange.Value;  // field names
                //fieldStrings = fieldRange.Value2;
                Excel.Worksheet sourceSheet = ThisAddIn.MyApp.get_Range((object)this.ContainingSheet.xlLinkCell.Value).Worksheet;
                
                for (int i = 1; i <= this.FieldCount; i++)
                {
                    fieldDict.Add(new UniqueID(sourceSheet.Name, fieldStrings[1,i].ToString()), i);      //is this being launched off correlation sheet? If so, have to follow the link
                }
                return fieldDict;
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
                return null; //No names in UniqueIDs anymore..
                //return FieldDict.Keys.Select(x => x.Name).ToArray<object>();
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

            public bool Equals(CorrelationMatrix cm)
            {
                if (cm.Matrix.GetLength(0) != this.Matrix.GetLength(0) || cm.Matrix.GetLength(1) != this.Matrix.GetLength(1))
                    return false;
                for(int row = 0; row < this.Matrix.GetLength(0); row++)
                {
                    for(int col = row; col < this.Matrix.GetLength(1); col++)
                    {
                        if (!Double.TryParse(this.Matrix[row, col].ToString(), out double internalVal))
                            throw new Exception("Invalid matrix value");
                        if (!Double.TryParse(cm.Matrix[row, col].ToString(), out double externalVal))
                            throw new Exception("Invalid matrix value");
                        if (internalVal != externalVal)
                            return false;
                    }
                }
                return true;
            }

            public bool ValidateAgainstTriple(PhasingTriple pt)
            {
                Data.CorrelationMatrix tripleMatrix = pt.GetPhasingCorrelationMatrix(this.FieldCount);
                return this.Equals(tripleMatrix);
            }
        }   //class
    }//Data
}//namespace
