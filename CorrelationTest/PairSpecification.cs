using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Accord.Statistics.Models.Regression.Linear;
using Excel = Microsoft.Office.Interop.Excel;

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

        public static PairSpecification ConstructFromString(string pairString)
        {
            PairSpecification pairSpec = new PairSpecification();
            //Header
            //Pair 1
            //Pair 2
            // ...
            //Pair N
            string[] lines = pairString.Split('&');
            pairSpec.Pairs = new Tuple<double, double>[lines.Count() - 1];
            for(int i = 1; i < lines.Length; i++)
            {
                string[] pair = lines[i].Split(',');
                pairSpec.Pairs[i - 1] = new Tuple<double, double>(Convert.ToDouble(pair[0]), Convert.ToDouble(pair[1]));
            }
            pairSpec.Value = pairSpec.ToString();
            return pairSpec;
        }

        public static PairSpecification ConstructByFittingMatrix(object[,] matrix, bool forceFitToOffDiagonal = false)
        {
            int size = matrix.GetLength(1);
            Tuple<double, double>[] pairs = new Tuple<double, double>[size-1];
            if(matrix.GetLength(0) != size)
            {
                throw new Exception("Matrix is not square");
            }
            for(int row = 0; row < size - 1; row++)
            {
                for (int col = row+2; col < size; col++)
                {
                    double[] points = GetPoints(row, matrix);
                    if (!forceFitToOffDiagonal)
                        pairs[row] = GetPair(points);
                    else
                        throw new NotImplementedException();
                }
            }
            PairSpecification pspec = new PairSpecification();
            pspec.Pairs = pairs;
            pspec.Value = pspec.ToString();
            return pspec;
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

        public object[,] GetCorrelationMatrix(string[] fields)
        {
            int size = fields.Count();
            object[,] matrix = new object[size, size];
            matrix[size - 1, size - 1] = 1;
            for(int row = 0; row < size-1; row++)
            {
                matrix[row, row] = 1;
                matrix[row, row + 1] = Pairs[row].Item1;
                
                for (int upIndex = 1; upIndex <= row; upIndex++)
                {
                    matrix[row - upIndex, row+1] = Pairs[row].Item1 - (Pairs[row].Item2 * (upIndex));
                }
                for(int downIndex = 1; downIndex < size - row; downIndex++)
                {
                    matrix[row + downIndex, row] = $"=OFFSET(INDIRECT(ADDRESS(ROW(),COLUMN(),4,1)),-{downIndex},{downIndex})";
                }
            }
            return matrix;
        }
    }
}
