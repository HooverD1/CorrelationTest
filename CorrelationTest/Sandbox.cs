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


    }
}
