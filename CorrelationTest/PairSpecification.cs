using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Accord.Statistics.Models.Regression.Linear;
using Excel = Microsoft.Office.Interop.Excel;
using System.Threading;

namespace CorrelationTest
{
    public class PairSpecification
    {
        private Tuple<double, double>[] Pairs { get; set; }      //<off diag value, vertical linear reduction>
        public string Value { get; set; }

        public static PairSpecification ConstructFromSinglePair(int matrixSize, double offDiagonal, double verticalDelta)
        {
            PairSpecification pairSpec = new PairSpecification();
            pairSpec.Pairs = new Tuple<double, double>[matrixSize-1];
            for (int i = 0; i < matrixSize-1; i++)
            {
                pairSpec.Pairs[i] = new Tuple<double, double>(offDiagonal, verticalDelta);
            }
            pairSpec.Value = pairSpec.ToString();
            return pairSpec;
        }

        //COLLAPSE
        public static PairSpecification ConstructFromRange(Excel.Range xlPairsCell, int sizeOfRange)    //Pull the pair spec from the correl sheet
        {
            StringBuilder sb = new StringBuilder();
            Excel.Range xlPairsRange = xlPairsCell.Resize[sizeOfRange, 2];
            for(int row = 1; row <= sizeOfRange; row++)
            {
                sb.Append((string)Convert.ToString(xlPairsRange.Cells[row, 1].value));
                sb.Append(",");
                sb.Append((string)Convert.ToString(xlPairsRange.Cells[row, 2].value));
                sb.Append("&");
            }
            sb.Remove(sb.Length - 1, 1);    //remove the final char
            return PairSpecification.ConstructFromString(sb.ToString(), false);
        }

        public static PairSpecification ConstructFromString(string pairString, bool includesHeader = true)
        {
            PairSpecification pairSpec = new PairSpecification();
            //Header
            //Pair 1
            //Pair 2
            // ...
            //Pair N
            if (includesHeader)
            {   //Remove the header if it is included
                pairString = pairString.Substring(pairString.IndexOf('&') + 1);
            }
            pairSpec.Value = pairString;
            string[] lines = pairString.Split('&');
            pairSpec.Pairs = new Tuple<double, double>[lines.Count()];
            for(int i = 0; i < lines.Length; i++)
            {
                string[] pair = lines[i].Split(',');
                pairSpec.Pairs[i] = new Tuple<double, double>(Convert.ToDouble(pair[0]), Convert.ToDouble(pair[1]));
            }
            
            return pairSpec;
        }

        private static double[] GetPoints(int row, object[,] matrix)
        {
            double[] points = new double[row];
            for(int i = 0; i < row; i++)
            {
                points[i] = Convert.ToDouble(matrix[i, row]);
            }
            return points;
        }

        private static Tuple<double, double> GetPair(double[] points)
        {
            double[] xValues = new double[points.Length];
            for (int i = 0; i < points.Length; i++)
                xValues[i] = i;
            var reg = SimpleLinearRegression.FromData(xValues, points);
            return new Tuple<double, double>(reg.Intercept, reg.Slope);
        }

        public override string ToString()
        {
            StringBuilder sb = new StringBuilder();
            for(int i = 0; i < Pairs.Length; i++)
            {
                sb.Append($"{Pairs[i].Item1},{Pairs[i].Item2}&");
            }
            sb.Remove(sb.Length - 1, 1);    //remove the final char
            return sb.ToString();
        }

        public double[,] GetCorrelationMatrix_Values()
        {
            int size = this.Pairs.Count() + 1;
            double[,] matrix = new double[size, size];
            matrix[size - 1, size - 1] = 1;
            for(int row = 0; row < size; row++)
            {
                matrix[row, row] = 1;
                //matrix[row, row + 1] = Pairs[row].Item1;
                
                for (int upIndex = 1; upIndex <= row; upIndex++)
                {
                    matrix[row - upIndex, row] = Pairs[row-1].Item1 - (Pairs[row-1].Item2 * (upIndex-1));
                }
            }
            for(int itRow = 1; itRow < size; itRow++)
            {
                for(int itCol = 0; itCol < itRow; itCol++)
                {
                    matrix[itRow, itCol] = matrix[itCol, itRow];
                }
            }
            return matrix;
        }

        //Serial version -- ~.1 second slower
        //public object[,] Old_GetCorrelationMatrix_Formulas(Sheets.CorrelationSheet CorrelSheet)
        //{
        //    Diagnostics.StartTimer();
        //    int size = this.Pairs.Count() + 1;
        //    Excel.Worksheet xlCorrelSheet = CorrelSheet.xlSheet;
        //    Data.CorrelSheetSpecs specs = CorrelSheet.Specs;
        //    Excel.Range pairsRange = CorrelSheet.xlPairsCell;
        //    int startRow = pairsRange.Row;
        //    int startCol = pairsRange.Column;
        //    object[,] matrix = new object[size, size];
        //    matrix[size - 1, size - 1] = "1";
        //    for (int row = 0; row < size - 1; row++)
        //    {
        //        matrix[row, row] = "1";
        //        matrix[row, row + 1] = $"=MIN(1,MAX(-1,R{startRow + row}C{startCol}))";

        //        for (int upIndex = 1; upIndex <= row; upIndex++)
        //        {
        //            matrix[row - upIndex, row + 1] = $"=MIN(1,MAX(-1,R{startRow + row}C{startCol} - R{startRow + row}C{startCol + 1} * {upIndex}))";
        //        }
        //        for (int downIndex = 1; downIndex < size - row; downIndex++)
        //        {
        //            matrix[row + downIndex, row] = $"=OFFSET(INDIRECT(ADDRESS(ROW(),COLUMN(),4,1)),-{downIndex},{downIndex})";
        //        }
        //    }
        //    long time = Diagnostics.CheckTimer();
        //    Diagnostics.StopTimer();
        //    return matrix;
        //}

        public string[,] GetCorrelationMatrix_Formulas(Sheets.CorrelationSheet CorrelSheet)
        {
            Diagnostics.StartTimer();
            int size = this.Pairs.Count() + 1;
            Excel.Worksheet xlCorrelSheet = CorrelSheet.xlSheet;
            Data.CorrelSheetSpecs specs = CorrelSheet.Specs;
            Excel.Range pairsRange = CorrelSheet.xlPairsCell;

            string[,] matrix = new string[size, size];
            string[,] addresses = new string[size, size];
            matrix[size - 1, size - 1] = "1";
            int startRow = pairsRange.Row;
            int startCol = pairsRange.Column;
            
            //This does the off-diagonal
            for (int row = 0; row < size - 1; row++)
            {
                string minBound = $"IF(R{startRow + row}C{startCol}>0,0,-1)";
                string maxBound = $"IF(R{startRow + row}C{startCol}<0,0,1)";
                matrix[row, row] = "=1";
                matrix[row, row + 1] = $"=MIN({maxBound},MAX({minBound},R{startRow + row}C{startCol}))";
            }
            void LoadUpperTriangular()
            {
                for (int row = 0; row < size - 1; row++)
                {
                    for (int rightIndex = 1; rightIndex <= size - row - 2; rightIndex++)        //Getting the .Address off the cell is slowing it down... and probably causing conflicts w threading
                    {
                        string minBound = $"IF(R{startRow + row}C{startCol}>0,0,-1)";
                        string maxBound = $"IF(R{startRow + row}C{startCol}<0,0,1)";
                        matrix[row, row + rightIndex + 1] = $"=MIN({maxBound},MAX({minBound},R{startRow + row}C{startCol} - R{startRow + row}C{startCol + 1} * {rightIndex}))";
                    }
                }
            }
            void LoadLowerTriangular()
            {
                for (int row = 0; row < size - 1; row++)
                {
                    for (int downIndex = 1; downIndex < size - row; downIndex++)
                    {
                        matrix[row + downIndex, row] = $"=OFFSET(INDIRECT(ADDRESS(ROW(),COLUMN(),4,1)),-{downIndex},{downIndex})";
                    }
                }
            }
            Thread th1 = new Thread(() => LoadUpperTriangular());
            th1.Start();
            Thread th2 = new Thread(() => LoadLowerTriangular());
            th2.Start();
            th1.Join();
            th2.Join();
            long time = Diagnostics.CheckTimer();
            Diagnostics.StopTimer();
            return matrix;
        }

        public object[,] GetValuesString_Split()
        {
            string[] lines = this.Value.Split('&');
            object[,] returnValues = new object[lines.Length, 2];
            for (int i = 0; i < lines.Length; i++)
            {
                string[] pair = lines[i].Split(',');
                returnValues[i, 0] = pair[0];
                returnValues[i, 1] = pair[1];
            }
            return returnValues;
        }

        private object[,] ConvertTuplesToRangeValues()
        {
            int numberOfPairs = this.Pairs.Count();
            object[,] rangeValues = new object[numberOfPairs, 2];
            for (int i=0;i<numberOfPairs; i++)
            {
                rangeValues[i, 0] = Pairs[i].Item1;
                rangeValues[i, 1] = Pairs[i].Item2;
            }
            return rangeValues;
        }

        public void PrintToSheet(Excel.Range xlPrintCell)
        {
            //Resize to fit the pairs
            Excel.Range xlPrintRange = xlPrintCell.Resize[this.Pairs.Count(), 2];
            //Convert tuples to object array
            xlPrintRange.Value = ConvertTuplesToRangeValues();
        }

        public static PairSpecification ConstructByFittingMatrix(object[,] matrixRange, bool forceFitDiagonal = false)
        {
            if (matrixRange == null)
                throw new ArgumentNullException();
            if (matrixRange.GetLength(0) != matrixRange.GetLength(1))
                throw new Exception("Matrix is not square");
            PairSpecification ps = new PairSpecification();
            //Give back an array of pairwise spec values
            Tuple<double, double>[] pairs = new Tuple<double, double>[matrixRange.GetLength(0) - 1];
            object[][] jaggedMatrix = ExtensionMethods.ToJaggedArray(matrixRange, true);

            for(int row = 1; row < matrixRange.GetLength(0); row++)
            {
                //yVals needs populated by the values above the i'th position
                double[] yVals = new double[matrixRange.GetLength(0) - row];
                double[] xVals = new double[matrixRange.GetLength(0) - row];
                for (int x = 0; x < matrixRange.GetLength(0) - row; x++)
                {
                    xVals[x] = row - x - 1;
                    yVals[x] = Convert.ToDouble(jaggedMatrix[row-1][row + x]);
                }
                    
                SimpleLinearRegression slr;
                var ols = new OrdinaryLeastSquares();
                double verticalShift = 0;
                if (forceFitDiagonal)
                {
                    ols.UseIntercept = false;
                    //Have to shift the y values down by fx(0) so that fx(0) = 0.
                    //Then run with .UseIntercept = false and add fx(0) to each slr.Intercept value
                    verticalShift = yVals[0];
                    if(verticalShift != 0)
                    {

                    }
                    for (int j = 0; j < yVals.Length; j++)
                    {
                        yVals[j] -= verticalShift;
                    }
                }
                if (row == matrixRange.GetLength(0) - 1)        //Last row has only the off diagonal, no reduction
                    pairs[row-1] = new Tuple<double, double>(yVals[0], 0);
                else
                {
                    try
                    {
                        slr = ols.Learn(xVals, yVals);
                        pairs[row-1] = new Tuple<double, double>(slr.Intercept + verticalShift, slr.Slope);  //Invert the slope because it stores as the "decrease"
                    }
                    catch
                    {
                        if (MyGlobals.DebugMode)
                            throw new Exception("OLS.learn failure");
                    }
                }
            }
            ps.Pairs = pairs;
            ps.Value = ps.ToString();
            return ps;
        }
    }
}
