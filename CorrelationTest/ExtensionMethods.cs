using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using Accord.Math.Decompositions;
using Accord.Math;

namespace CorrelationTest
{
    public static class ExtensionMethods
    {
        internal static T[][] ToJaggedArray<T>(this T[,] twoDimensionalArray, bool transpose = false)
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
                    if(!transpose)
                        jaggedArray[i][j] = twoDimensionalArray[i, j];
                    else
                        jaggedArray[i][j] = twoDimensionalArray[j, i];
                }
            }
            return jaggedArray;
        }

        public static T[,] Transpose<T>(T[,] inputArray)
        {
            int size = inputArray.GetLength(0);
            T[,] returnArray = new T[size, size];
            for(int r = 0; r < size; r++)
            {
                for(int c = 0; c < size; c++)
                {
                    returnArray[c, r] = inputArray[r, c];
                }
            }
            return returnArray;
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
                    return "$CORRELATION_CP";
                case SheetType.Correlation_CM:
                    return "$CORRELATION_CM";
                case SheetType.Correlation_PM:
                    return "$CORRELATION_PM";
                case SheetType.Correlation_PP:
                    return "$CORRELATION_PP";
                case SheetType.Correlation_DP:
                    return "$CORRELATION_DP";
                case SheetType.Correlation_DM:
                    return "$CORRELATION_DM";
                default:
                    return null;
            }
        }

        public static SheetType GetSheetType(Excel.Worksheet xlSheet)
        {
            string sheetIdent = Convert.ToString(xlSheet.Cells[1, 1].Value);
            switch (sheetIdent)
            {
                case "$CORRELATION_CP":
                    return SheetType.Correlation_CP;
                case "$CORRELATION_CM":
                    return SheetType.Correlation_CM;
                case "$CORRELATION_PM":
                    return SheetType.Correlation_PM;
                case "$CORRELATION_PP":
                    return SheetType.Correlation_PP;
                case "$CORRELATION_DM":
                    return SheetType.Correlation_DM;
                case "$CORRELATION_DP":
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

        public static object[,] Convert1Dto2D(object[] inputArray)
        {
            object[,] verticalArray = new object[inputArray.Length, 1];
            for(int i = 0; i < inputArray.Length; i++)
            {
                verticalArray[i, 0] = inputArray[i];
            }
            return verticalArray;
        }

        public static CorrelationType GetCorrelationTypeFromLink(Excel.Range linkSource)
        {
            SheetType sheetType = ExtensionMethods.GetSheetType(linkSource.Worksheet);
            DisplayCoords dc = DisplayCoords.ConstructDisplayCoords(sheetType);
            string parentType = Convert.ToString(linkSource.EntireRow.Cells[1, dc.Type_Offset].value);

            CorrelationType cType = GetCorrelationTypeFromItemType(parentType);
            if (cType == CorrelationType.Null)
            {
                if (dc.PhasingCorrel_Offset == linkSource.Column)
                {
                    return CorrelationType.Phasing;
                }
                else
                {
                    throw new Exception("Unknown correlation type");
                }
            }
            else
            {
                return cType;
            }
            
        }
        public static CorrelationType GetCorrelationTypeFromItemType(string itemType)
        {
            switch (itemType)
            {
                case "CASE":
                    return CorrelationType.Cost;
                case "SACE":
                    return CorrelationType.Duration;
                case "CE":
                    return CorrelationType.Cost;
                case "SE":
                    return CorrelationType.Duration;
                default:
                    return CorrelationType.Null;
            }
        }
        public static string[] GetStringFromObject(object[] inObject)
        {
            if (inObject == null)
                throw new Exception("Object is null");
            if(inObject.GetLowerBound(0) == 1)
                inObject = ReIndexArray(inObject);
            string[] outString = new string[inObject.Length];
            for(int i = 0; i < inObject.Length; i++)
            {
                outString[i] = Convert.ToString(inObject[i]);
            }
            return outString;
        }

               
        //public static void SturdyPaste_Square(Excel.Range pasteRange, object[,] pasteValues)
        //{
        //    //Error checking
        //    int size = pasteRange.Columns.Count;
        //    if (size != pasteRange.Rows.Count)
        //        throw new Exception("Not square");
        //    if (pasteValues.GetLength(0) != size || pasteValues.GetLength(1) != size)
        //        throw new Exception("Values array size does not match range size");
        //    //Error checking

        //    //This method splits up the inputs into chunks of max size 250 x 250 and handles them separately

        //    Excel.Range pasteCell = pasteRange.Cells[1, 1];     //The top left cell
        //    Excel.Range[] partialPasteRange;     //Max 250 x 250 cell blocks
        //    object[] partialValues;

        //    void ThreadMe(int[,] parameters)
        //    {
        //        for(int block = 0; block < parameters.GetLength(1); block++)
        //        {
        //            ProcessBlock(parameters[4,block], parameters[0,block], parameters[1,block], parameters[2,block], parameters[3,block]);
        //        }
        //    }

        //    void ProcessBlock(int blockIndex, int param1, int param2, int param3, int param4)
        //    {
        //        //partialPasteRange[blockIndex] = pasteCell.Offset[param1, param2].Resize[param3, param4];
        //        partialValues[blockIndex] = SliceArray(pasteValues, param1, param2, param3, param4);
        //        partialPasteRange[blockIndex].Value = partialValues[blockIndex];
        //    }

        //    if (size <= 250)
        //    {
        //        partialPasteRange = new Excel.Range[1];
        //        partialPasteRange[0] = pasteCell.Offset[0, 0].Resize[size, size];

        //        partialValues = new object[1];
        //        partialValues[0] = pasteValues;     //object arrays are themselves objects..
        //        partialPasteRange[0].FormulaR1C1 = partialValues[0];
        //    }
        //    else if(size <= 500)
        //    {               
        //        //2 x 2 blocks of roughly 250 x 250
        //        partialPasteRange = new Excel.Range[4];
        //        int blockSize = size / 2 + (size % 2)/2;
        //        partialPasteRange[0] = pasteCell.Offset[0, 0].Resize[blockSize, blockSize];     //full width
        //        partialPasteRange[1] = pasteCell.Offset[0, blockSize].Resize[blockSize, size - blockSize]; //partial width
        //        partialPasteRange[2] = pasteCell.Offset[blockSize, 0].Resize[size - blockSize, blockSize]; //full width
        //        partialPasteRange[3] = pasteCell.Offset[blockSize, blockSize].Resize[size - blockSize, size - blockSize]; //partial width

        //        partialValues = new object[4];

        //        //Thread th1 = new Thread(() => ProcessBlock(0, 0, 0, blockSize, blockSize));
        //        //Thread th2 = new Thread(() => ProcessBlock(1, 0, blockSize, blockSize, size - blockSize));
        //        //Thread th3 = new Thread(() => ProcessBlock(2, blockSize, 0, size - blockSize, blockSize));
        //        //Thread th4 = new Thread(() => ProcessBlock(3, blockSize, blockSize, size - blockSize, size - blockSize));
        //        //th1.Start();
        //        //th2.Start();
        //        //th3.Start();
        //        //th4.Start();
        //        //th1.Join();
        //        //th2.Join();
        //        //th3.Join();
        //        //th4.Join();

        //        partialValues[0] = SliceArray(pasteValues, 0, 0, blockSize, blockSize);
        //        partialValues[1] = SliceArray(pasteValues, 0, blockSize, blockSize, size - blockSize);
        //        partialValues[2] = SliceArray(pasteValues, blockSize, 0, size - blockSize, blockSize);
        //        partialValues[3] = SliceArray(pasteValues, blockSize, blockSize, size - blockSize, size - blockSize);

        //        partialPasteRange[0].FormulaR1C1 = partialValues[0];      //Can I multi-thread this?
        //        partialPasteRange[1].FormulaR1C1 = partialValues[1];
        //        partialPasteRange[2].FormulaR1C1 = partialValues[2];
        //        partialPasteRange[3].FormulaR1C1 = partialValues[3];
        //    }
        //    else if(size <= 750)
        //    {
        //        //3 x 3 blocks of 250 x 250
        //        partialPasteRange = new Excel.Range[9];
        //        int blockSize = size / 3 + (size % 3) / 3;
        //        partialPasteRange[0] = pasteCell.Offset[blockSize * 0, blockSize*0].Resize[blockSize, blockSize];   //full width
        //        partialPasteRange[1] = pasteCell.Offset[blockSize * 0, blockSize*1].Resize[blockSize, blockSize];   //full width
        //        partialPasteRange[2] = pasteCell.Offset[blockSize * 0, blockSize*2].Resize[blockSize, size - (blockSize*2)];   //partial width
        //        partialPasteRange[3] = pasteCell.Offset[blockSize * 1, blockSize*0].Resize[blockSize, blockSize];   //full width
        //        partialPasteRange[4] = pasteCell.Offset[blockSize * 1, blockSize*1].Resize[blockSize, blockSize];   //full width
        //        partialPasteRange[5] = pasteCell.Offset[blockSize * 1, blockSize*2].Resize[blockSize, size - (blockSize * 2)];   //partial width
        //        partialPasteRange[6] = pasteCell.Offset[blockSize * 2, blockSize*0].Resize[size - (blockSize * 2), blockSize];   //full width
        //        partialPasteRange[7] = pasteCell.Offset[blockSize * 2, blockSize*1].Resize[size - (blockSize * 2), blockSize];   //full width
        //        partialPasteRange[8] = pasteCell.Offset[blockSize * 2, blockSize*2].Resize[size - (blockSize * 2), size - (blockSize * 2)];   //partial width

        //        partialValues = new object[9];
        //        //Thread th1 = new Thread(() => ProcessBlock(0, blockSize * 0, blockSize * 0, blockSize, blockSize));
        //        //th1.Start();
        //        //Thread th2 = new Thread(() => ProcessBlock(1, blockSize * 0, blockSize * 1, blockSize, blockSize));
        //        //th2.Start();
        //        //Thread th3 = new Thread(() => ProcessBlock(2, blockSize * 0, blockSize * 2, blockSize, size - blockSize * 2));
        //        //th3.Start();
        //        //th1.Join();
        //        //th1 = new Thread(() => ProcessBlock(3, blockSize * 1, blockSize * 0, blockSize, blockSize));
        //        //th1.Start();
        //        //th2.Join();
        //        //th2 = new Thread(() => ProcessBlock(4, blockSize * 1, blockSize * 1, blockSize, blockSize));
        //        //th2.Start();
        //        //th3.Join();
        //        //th3 = new Thread(() => ProcessBlock(5, blockSize * 1, blockSize * 2, blockSize, size - blockSize * 2));
        //        //th3.Start();
        //        //th1.Join();
        //        //th1 = new Thread(() => ProcessBlock(6, blockSize * 2, blockSize * 0, blockSize, blockSize));
        //        //th1.Start();
        //        //th2.Join();
        //        //th2 = new Thread(() => ProcessBlock(7, blockSize * 2, blockSize * 1, blockSize, blockSize));
        //        //th2.Start();
        //        //th3.Join();
        //        //th3 = new Thread(() => ProcessBlock(8, blockSize * 2, blockSize * 2, blockSize, size - blockSize * 2));
        //        //th3.Start();
        //        //th1.Join();
        //        //th2.Join();
        //        //th3.Join();

        //        partialValues[0] = SliceArray(pasteValues, blockSize * 0, blockSize * 0, blockSize, blockSize);
        //        partialValues[1] = SliceArray(pasteValues, blockSize * 0, blockSize * 1, blockSize, blockSize);
        //        partialValues[2] = SliceArray(pasteValues, blockSize * 0, blockSize * 2, blockSize, size - blockSize * 2);
        //        partialValues[3] = SliceArray(pasteValues, blockSize * 1, blockSize * 0, blockSize, blockSize);
        //        partialValues[4] = SliceArray(pasteValues, blockSize * 1, blockSize * 1, blockSize, blockSize);
        //        partialValues[5] = SliceArray(pasteValues, blockSize * 1, blockSize * 2, blockSize, size - blockSize * 2);
        //        partialValues[6] = SliceArray(pasteValues, blockSize * 2, blockSize * 0, blockSize, blockSize);
        //        partialValues[7] = SliceArray(pasteValues, blockSize * 2, blockSize * 1, blockSize, blockSize);
        //        partialValues[8] = SliceArray(pasteValues, blockSize * 2, blockSize * 2, blockSize, size - blockSize * 2);

        //        partialPasteRange[0].FormulaR1C1 = partialValues[0];      //Can I multi-thread this?
        //        partialPasteRange[1].FormulaR1C1 = partialValues[1];
        //        partialPasteRange[2].FormulaR1C1 = partialValues[2];
        //        partialPasteRange[3].FormulaR1C1 = partialValues[3];
        //        partialPasteRange[4].FormulaR1C1 = partialValues[4];
        //        partialPasteRange[5].FormulaR1C1 = partialValues[5];
        //        partialPasteRange[6].FormulaR1C1 = partialValues[6];
        //        partialPasteRange[7].FormulaR1C1 = partialValues[7];
        //        partialPasteRange[8].FormulaR1C1 = partialValues[8];
        //    }
        //    else if(size <= 1000)
        //    {
        //        //4 x 4 blocks of 250 x 250
        //        partialPasteRange = new Excel.Range[16];
        //        int blockSize = size / 4 + (size % 4) / 4;
        //        partialPasteRange[0] = pasteCell.Offset[blockSize * 0, blockSize * 0].Resize[blockSize, blockSize];
        //        partialPasteRange[1] = pasteCell.Offset[blockSize * 0, blockSize * 1].Resize[blockSize, blockSize];
        //        partialPasteRange[2] = pasteCell.Offset[blockSize * 0, blockSize * 2].Resize[blockSize, blockSize];
        //        partialPasteRange[3] = pasteCell.Offset[blockSize * 0, blockSize * 3].Resize[blockSize, size - blockSize*3];//
        //        partialPasteRange[4] = pasteCell.Offset[blockSize * 1, blockSize * 0].Resize[blockSize, blockSize];
        //        partialPasteRange[5] = pasteCell.Offset[blockSize * 1, blockSize * 1].Resize[blockSize, blockSize];
        //        partialPasteRange[6] = pasteCell.Offset[blockSize * 1, blockSize * 2].Resize[blockSize, blockSize];
        //        partialPasteRange[7] = pasteCell.Offset[blockSize * 1, blockSize * 3].Resize[blockSize, size - blockSize * 3];//
        //        partialPasteRange[8] = pasteCell.Offset[blockSize * 2, blockSize * 0].Resize[blockSize, blockSize];
        //        partialPasteRange[9] = pasteCell.Offset[blockSize * 2, blockSize * 1].Resize[blockSize, blockSize];
        //        partialPasteRange[10] = pasteCell.Offset[blockSize * 2, blockSize * 2].Resize[blockSize, blockSize];
        //        partialPasteRange[11] = pasteCell.Offset[blockSize * 2, blockSize * 3].Resize[blockSize, size - blockSize * 3];//
        //        partialPasteRange[12] = pasteCell.Offset[blockSize * 3, blockSize * 0].Resize[size - blockSize * 3, blockSize];
        //        partialPasteRange[13] = pasteCell.Offset[blockSize * 3, blockSize * 1].Resize[size - blockSize * 3, blockSize];
        //        partialPasteRange[14] = pasteCell.Offset[blockSize * 3, blockSize * 2].Resize[size - blockSize * 3, blockSize];
        //        partialPasteRange[15] = pasteCell.Offset[blockSize * 3, blockSize * 3].Resize[size - blockSize * 3, size - blockSize * 3];//

        //        partialValues = new object[16];
        //        //int[,] parameters = new int[5,4];       //[param, block]        //One of these per thread
        //        //parameters[0, 0] = blockSize * 0;
        //        //parameters[1, 0] = blockSize * 0;
        //        //parameters[2, 0] = blockSize;
        //        //parameters[3, 0] = blockSize;
        //        //parameters[4, 0] = 0;

        //        //parameters[0, 1] = blockSize * 0;
        //        //parameters[1, 1] = blockSize * 1;
        //        //parameters[2, 1] = blockSize;
        //        //parameters[3, 1] = blockSize;
        //        //parameters[4, 1] = 1;

        //        //parameters[0, 2] = blockSize * 0;
        //        //parameters[1, 2] = blockSize * 2;
        //        //parameters[2, 2] = blockSize;
        //        //parameters[3, 2] = blockSize;
        //        //parameters[4, 2] = 2;

        //        //parameters[0, 3] = blockSize * 0;
        //        //parameters[1, 3] = blockSize * 3;
        //        //parameters[2, 3] = blockSize;
        //        //parameters[3, 3] = size - blockSize * 3;
        //        //parameters[4, 3] = 3;

        //        ////oh christ
        //        //Thread th1 = new Thread(() => ThreadMe(parameters));
        //        //th1.Start();

        //        //parameters[0, 0] = blockSize * 1;
        //        //parameters[1, 0] = blockSize * 0;
        //        //parameters[2, 0] = blockSize;
        //        //parameters[3, 0] = blockSize;
        //        //parameters[4, 0] = 4;

        //        //parameters[0, 1] = blockSize * 1;
        //        //parameters[1, 1] = blockSize * 1;
        //        //parameters[2, 1] = blockSize;
        //        //parameters[3, 1] = blockSize;
        //        //parameters[4, 1] = 5;

        //        //parameters[0, 2] = blockSize * 1;
        //        //parameters[1, 2] = blockSize * 2;
        //        //parameters[2, 2] = blockSize;
        //        //parameters[3, 2] = blockSize;
        //        //parameters[4, 2] = 6;

        //        //parameters[0, 3] = blockSize * 1;
        //        //parameters[1, 3] = blockSize * 3;
        //        //parameters[2, 3] = blockSize;
        //        //parameters[3, 3] = size - blockSize * 3;
        //        //parameters[4, 3] = 7;

        //        //Thread th2 = new Thread(() => ThreadMe(parameters));
        //        //th2.Start();

        //        //parameters[0, 0] = blockSize * 2;
        //        //parameters[1, 0] = blockSize * 0;
        //        //parameters[2, 0] = blockSize;
        //        //parameters[3, 0] = blockSize;
        //        //parameters[4, 0] = 8;

        //        //parameters[0, 1] = blockSize * 2;
        //        //parameters[1, 1] = blockSize * 1;
        //        //parameters[2, 1] = blockSize;
        //        //parameters[3, 1] = blockSize;
        //        //parameters[4, 1] = 9;

        //        //parameters[0, 2] = blockSize * 2;
        //        //parameters[1, 2] = blockSize * 2;
        //        //parameters[2, 2] = blockSize;
        //        //parameters[3, 2] = blockSize;
        //        //parameters[4, 2] = 10;

        //        //parameters[0, 3] = blockSize * 2;
        //        //parameters[1, 3] = blockSize * 3;
        //        //parameters[2, 3] = blockSize;
        //        //parameters[3, 3] = size - blockSize * 3;
        //        //parameters[4, 3] = 11;

        //        //Thread th3 = new Thread(() => ThreadMe(parameters));
        //        //th3.Start();

        //        //parameters[0, 0] = blockSize * 3;
        //        //parameters[1, 0] = blockSize * 0;
        //        //parameters[2, 0] = size - blockSize * 3;
        //        //parameters[3, 0] = blockSize;
        //        //parameters[4, 0] = 12;

        //        //parameters[0, 1] = blockSize * 3;
        //        //parameters[1, 1] = blockSize * 1;
        //        //parameters[2, 1] = size - blockSize * 3;
        //        //parameters[3, 1] = blockSize;
        //        //parameters[4, 1] = 13;

        //        //parameters[0, 2] = blockSize * 3;
        //        //parameters[1, 2] = blockSize * 2;
        //        //parameters[2, 2] = size - blockSize * 3;
        //        //parameters[3, 2] = blockSize;
        //        //parameters[4, 2] = 14;

        //        //parameters[0, 3] = blockSize * 3;
        //        //parameters[1, 3] = blockSize * 3;
        //        //parameters[2, 3] = size - blockSize * 3;
        //        //parameters[3, 3] = size - blockSize * 3;
        //        //parameters[4, 3] = 15;

        //        //Thread th4 = new Thread(() => ThreadMe(parameters));
        //        //th4.Start();

        //        //th1.Join();
        //        //th2.Join();
        //        //th3.Join();
        //        //th4.Join();

        //        partialValues[0] = SliceArray(pasteValues, blockSize * 0, blockSize * 0, blockSize, blockSize);
        //        partialValues[1] = SliceArray(pasteValues, blockSize * 0, blockSize * 1, blockSize, blockSize);
        //        partialValues[2] = SliceArray(pasteValues, blockSize * 0, blockSize * 2, blockSize, blockSize);
        //        partialValues[3] = SliceArray(pasteValues, blockSize * 0, blockSize * 3, blockSize, size - blockSize * 3);
        //        partialValues[4] = SliceArray(pasteValues, blockSize * 1, blockSize * 0, blockSize, blockSize);
        //        partialValues[5] = SliceArray(pasteValues, blockSize * 1, blockSize * 1, blockSize, blockSize);
        //        partialValues[6] = SliceArray(pasteValues, blockSize * 1, blockSize * 2, blockSize, blockSize);
        //        partialValues[7] = SliceArray(pasteValues, blockSize * 1, blockSize * 3, blockSize, size - blockSize * 3);
        //        partialValues[8] = SliceArray(pasteValues, blockSize * 2, blockSize * 0, blockSize, blockSize);
        //        partialValues[9] = SliceArray(pasteValues, blockSize * 2, blockSize * 1, blockSize, blockSize);
        //        partialValues[10] = SliceArray(pasteValues, blockSize * 2, blockSize * 2, blockSize, blockSize);
        //        partialValues[11] = SliceArray(pasteValues, blockSize * 2, blockSize * 3, blockSize, size - blockSize * 3);
        //        partialValues[12] = SliceArray(pasteValues, blockSize * 3, blockSize * 0, size - blockSize * 3, blockSize);
        //        partialValues[13] = SliceArray(pasteValues, blockSize * 3, blockSize * 1, size - blockSize * 3, blockSize);
        //        partialValues[14] = SliceArray(pasteValues, blockSize * 3, blockSize * 2, size - blockSize * 3, blockSize);
        //        partialValues[15] = SliceArray(pasteValues, blockSize * 3, blockSize * 3, size - blockSize * 3, size - blockSize * 3);

        //        partialPasteRange[0].FormulaR1C1 = partialValues[0];      //Can I multi-thread this?
        //        partialPasteRange[1].FormulaR1C1 = partialValues[1];
        //        partialPasteRange[2].FormulaR1C1 = partialValues[2];
        //        partialPasteRange[3].FormulaR1C1 = partialValues[3];
        //        partialPasteRange[4].FormulaR1C1 = partialValues[4];
        //        partialPasteRange[5].FormulaR1C1 = partialValues[5];
        //        partialPasteRange[6].FormulaR1C1 = partialValues[6];
        //        partialPasteRange[7].FormulaR1C1 = partialValues[7];
        //        partialPasteRange[8].FormulaR1C1 = partialValues[8];
        //        partialPasteRange[9].FormulaR1C1 = partialValues[9];
        //        partialPasteRange[10].FormulaR1C1 = partialValues[10];
        //        partialPasteRange[11].FormulaR1C1 = partialValues[11];
        //        partialPasteRange[12].FormulaR1C1 = partialValues[12];
        //        partialPasteRange[13].FormulaR1C1 = partialValues[13];
        //        partialPasteRange[14].FormulaR1C1 = partialValues[14];
        //        partialPasteRange[15].FormulaR1C1 = partialValues[15];
        //    }
        //    else
        //    {
        //        throw new Exception("Matrix is too larger");
        //    }
        //}

        public static void SturdyPaste2(Excel.Range pasteRange, object[][] pasteValues)
        {
            ThisAddIn.MyApp.ReferenceStyle = Excel.XlReferenceStyle.xlR1C1;
            for(int i = 1; i <= pasteRange.Rows.Count; i++)
            {
                pasteRange.Rows[i].FormulaR1C1 = pasteValues[i-1];
            }
            
        }

        public static string[] SliceArray(string[,] fullArray, int row_start, int columns)   //for generic object[x, y]
        {
            string[] slicedArray = new string[columns];
            for(int y = 0; y < columns; y++)
            {
                slicedArray[y] = fullArray[row_start, y];        //Does a quicker way than iteration exist?
            }
            return slicedArray;
        }

        public static T[,] ConvertObjectArray<T>(object[,] inputArray)      //Works for object --> double.. not tested on much else
        {
            inputArray = ExtensionMethods.ReIndexArray(inputArray);
            T[,] returnArray = new T[inputArray.GetLength(0), inputArray.GetLength(1)];
            for(int r = 0; r<inputArray.GetLength(0); r++)
            {
                for(int c=0; c<inputArray.GetLength(1); c++)
                {
                    try { returnArray[r, c] = (T)inputArray[r, c]; }
                    catch { throw new Exception("Failed conversion"); }
                }
            }
            return returnArray;
        }

        public static double CalculateStandardDeviation(IEnumerable<double> values)
        {
            double standardDeviation = 0;

            if (values.Any())
            {
                // Compute the average.     
                double avg = values.Average();

                // Perform the Sum of (value-avg)_2_2.      
                double sum = values.Sum(d => Math.Pow(d - avg, 2));

                // Put it all together.      
                standardDeviation = Math.Sqrt((sum) / (values.Count() - 1));
            }

            return standardDeviation;
        }

        public static void TurnOffUpdating()
        {
            Excel.Application App = ThisAddIn.MyApp;
            App.Calculation = Excel.XlCalculation.xlCalculationManual;
            App.ScreenUpdating = false;
            App.EnableEvents = false;
            App.DisplayStatusBar = false;
        }

        public static void TurnOnUpdating()
        {
            Excel.Application App = ThisAddIn.MyApp;
            App.ScreenUpdating = true;
            App.EnableEvents = true;
            App.DisplayStatusBar = true;
            App.Calculation = Excel.XlCalculation.xlCalculationAutomatic;
        }

        public static double[,] GetCholeskyDecomposition(double[,] matrix)
        {
            //Matrix must be symmetrical and PSD
            CholeskyDecomposition cholesky = new CholeskyDecomposition(matrix);
            return cholesky.DiagonalMatrix;
        }

        public static double GetStandardDeviation(double[] data)
        {
            double mean = data.Average();
            double variance = (from double d in data select Math.Pow(d - mean, 2)).Sum() / data.Length;
            return Math.Sqrt(variance);
        }
    }
}
