using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using Accord.Statistics.Models.Regression.Linear;
using System.Diagnostics;
using System.Windows.Forms;

namespace CorrelationTest
{
    public static class Sandbox
    {
        public static object[,] CreateRandomTestCorrelationMatrix(int size)
        {
            Random rando = new Random();
            object[,] testMatrix = new object[size, size];
            for(int row = 0; row < size; row++)
            {
                for (int col = row; col < size; col++)
                {
                    testMatrix[row, col] = rando.NextDouble() * 2 - 1;      // -1 to 1
                }
            }
            return testMatrix;
        }

        public static Tuple<double, double>[] FitMatrix(object[,] matrixRange, bool forceFitDiagonal = false)
        {
            Stopwatch sw = new Stopwatch();
            sw.Reset();
            sw.Start();
            //Give back an array of pairwise spec values
            Tuple<double, double>[] pairs = new Tuple<double, double>[matrixRange.GetLength(0) - 1];
            object[][] jaggedMatrix = ExtensionMethods.ToJaggedArray(matrixRange, true);
            
            for(int i = matrixRange.GetLength(0)-2; i >= 0; i--)
            {
                double[] yVals = (from object val in jaggedMatrix[i] where val != null select Convert.ToDouble(val)).ToArray();
                double[] xVals = new double[yVals.Length];
                for (int x = 0; x < yVals.Length; x++)
                    xVals[x] = yVals.Length - x - 1;
                SimpleLinearRegression slr;
                var ols = new OrdinaryLeastSquares();
                double verticalShift = 0;
                if (forceFitDiagonal)
                {
                    ols.UseIntercept = false;
                    //Have to shift the y values down by fx(0) so that fx(0) = 0.
                    //Then run with .UseIntercept = false and add fx(0) to each slr.Intercept value
                    verticalShift = yVals[yVals.Length - 1];
                    for(int j = 0; j < yVals.Length; j++)
                    {
                        yVals[i] -= verticalShift;
                    }
                }
                if (xVals.Length != yVals.Length)
                    throw new Exception("Malformed regression inputs");
                else if (xVals.Length < 2)
                    pairs[i] = new Tuple<double, double>(0, yVals[yVals.Length - 1]);
                else
                {
                    try
                    {
                        slr = ols.Learn(xVals, yVals);
                        pairs[i] = new Tuple<double, double>(slr.Slope, slr.Intercept + verticalShift);
                    }
                    catch
                    {
                        if(MyGlobals.DebugMode)
                            throw new Exception("OLS.learn failure");
                    }
                }
            }
            sw.Stop();
            MessageBox.Show(sw.ElapsedMilliseconds.ToString());
            return pairs;            
        }
    }
}
