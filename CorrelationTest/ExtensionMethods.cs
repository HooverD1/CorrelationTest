using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

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

        private static string TypeToLabel(SheetType sheetType)
        {
            switch (sheetType)
            {
                case SheetType.Estimate:
                    return "$EST";
                case SheetType.WBS:
                    return "$WBS";
                case SheetType.Data:
                    return "$DATA";
                case SheetType.Model:
                    return "$MODEL";
                case SheetType.Input:
                    return "$INPUT";
                case SheetType.FilterData:
                    return "$FILTER";
                case SheetType.Correlation_CP:
                    return "$CORRELATION_IT";
                case SheetType.Correlation_CM:
                    return "$CORRELATION_IM";
                case SheetType.Correlation_PM:
                    return "$CORRELATION_PM";
                case SheetType.Correlation_PP:
                    return "$CORRELATION_PT";
                case SheetType.Correlation_DP:
                    return "$CORRELATION_DT";
                case SheetType.Correlation_DM:
                    return "$CORRELATION_DM";
                default:
                    return null;
            }
        }

        public static SheetType GetSheetType(Excel.Worksheet xlSheet)
        {
            string sheetIdent = xlSheet.Cells[1, 1].Value;
            switch (sheetIdent)
            {
                case "$CORRELATION_CT":
                    return SheetType.Correlation_CP;
                case "$CORRELATION_CM":
                    return SheetType.Correlation_CM;
                case "$CORRELATION_PM":
                    return SheetType.Correlation_PM;
                case "$CORRELATION_PT":
                    return SheetType.Correlation_PP;
                case "$CORRELATION_DM":
                    return SheetType.Correlation_DM;
                case "$CORRELATION_DT":
                    return SheetType.Correlation_DP;
                case "$WBS":
                    return SheetType.WBS;
                case "$EST":
                    return SheetType.Estimate;
                default:
                    return SheetType.Unknown;
            }
        }

        public static object[,] AddLowerTriangular(object[,] upperTriangular)
        {
            //upperTriangular should be zero-based
            //upperTriangular = ReIndexArray<object>(upperTriangular);
            if (upperTriangular.GetLength(0) != upperTriangular.GetLength(1))
                throw new Exception("Correlation array not square");
            for(int row = 1; row < upperTriangular.GetLength(0); row++)
            {
                for(int col=0;col < row; col++)
                {
                    upperTriangular[row, col] = upperTriangular[col, row];
                }
            }
            return upperTriangular;
        }

        public static Excel.Worksheet GetWorksheet(string sheetName, SheetType sheetType = SheetType.Unknown)
        {
            Excel.Worksheet xlSheet;

            IEnumerable<Excel.Worksheet> xlSheets = from Excel.Worksheet sheet in ThisAddIn.MyApp.Worksheets
                                                    where sheet.Name == sheetName && sheet.Cells[1, 1].value == TypeToLabel(sheetType)
                                                    select sheet;
            if (xlSheets.Any())
            {
                xlSheet = xlSheets.First();
            }
            else
            {
                xlSheet = ThisAddIn.MyApp.Worksheets.Add();
                xlSheet.Name = sheetName;
                xlSheet.Cells[1, 1].value = TypeToLabel(sheetType);
            }
            return xlSheet;
        }

        public static object[,] GetSubArray(object[][] mainArray, int startIndex)
        {
            int matrixSize = mainArray[1].GetLength(0);
            object[,] returnArray = new object[matrixSize, matrixSize];
            for(int row = startIndex; row < mainArray.GetLength(0); row++)
            {
                for(int col = 0; col < mainArray[row].Length; col++)
                {
                    returnArray[row - startIndex, col] = mainArray[row][col];
                }
            }
            return returnArray;
        }

        public static string CleanStringLinebreaks(string my_string)
        {
            my_string = my_string.Replace("\r\n", "&");  //simplify delimiter
            my_string = my_string.Replace("\n", "&");  //simplify delimiter
            return my_string;
        }

        public static CorrelationType GetCorrelationTypeFromLink(Excel.Range linkSource)
        {
            SheetType sheetType = ExtensionMethods.GetSheetType(linkSource.Worksheet);
            DisplayCoords dc = DisplayCoords.ConstructDisplayCoords(sheetType);

            if(dc.CostCorrel_Offset == linkSource.Column)
            {
                return CorrelationType.Cost;
            }
            else if(dc.PhasingCorrel_Offset == linkSource.Column)
            {
                return CorrelationType.Phasing;
            }
            else if(dc.DurationCorrel_Offset == linkSource.Column)
            {
                return CorrelationType.Duration;
            }
            else
            {
                return CorrelationType.Null;    
            }
        }
    }
}
