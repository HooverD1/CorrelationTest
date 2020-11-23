using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CorrelationTest
{
    public static class ExtensionMethods
    {
        internal static T[][] ToJaggedArray<T>(this T[,] twoDimensionalArray)
        {
            int rowsFirstIndex = twoDimensionalArray.GetLowerBound(0);
            int rowsLastIndex = twoDimensionalArray.GetUpperBound(0);
            int numberOfRows = rowsLastIndex + 1;

            int columnsFirstIndex = twoDimensionalArray.GetLowerBound(1);
            int columnsLastIndex = twoDimensionalArray.GetUpperBound(1);
            int numberOfColumns = columnsLastIndex + 1;

            T[][] jaggedArray = new T[numberOfRows][];
            for (int i = rowsFirstIndex; i <= rowsLastIndex; i++)
            {
                jaggedArray[i] = new T[numberOfColumns];

                for (int j = columnsFirstIndex; j <= columnsLastIndex; j++)
                {
                    jaggedArray[i][j] = twoDimensionalArray[i, j];
                }
            }
            return jaggedArray;
        }

        public static T[] ReIndexArray<T>(T[] inputArray)
        {
            if (inputArray.GetLowerBound(0) > 0)
            {
                T[] copyArray = new T[inputArray.Length];
                Array.Copy(inputArray, copyArray, inputArray.Length);
                return copyArray;
            }
            else
                return inputArray;
        }
        public static T[,] ReIndexArray<T>(T[,] inputArray)
        {
            if (inputArray.GetLowerBound(0) > 0)
            {
                T[,] copyArray = new T[inputArray.GetLength(0), inputArray.GetLength(1)];
                Array.Copy(inputArray, copyArray, inputArray.Length);
                return copyArray;
            }
            else
                return inputArray;
        }
    }
}
