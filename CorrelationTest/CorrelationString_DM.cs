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
        public class CorrelationString_DM : CorrelationString
        {

            //EXPAND
            public CorrelationString_DM(string correlStringValue)
            {
                this.Value = correlStringValue;
            }

            public static CorrelationString_DM ConstructZeroString(string[] fields)
            {
                //Need to downcast csi 
                var csi = new CorrelationString(fields);
                return new CorrelationString_DM(csi.Value);
            }


            //COLLAPSE
            public CorrelationString_DM(Sheets.CorrelationSheet_DM correlSheet)
            {
                StringBuilder header = new StringBuilder();
                StringBuilder values = new StringBuilder();

                Excel.Range parentRow = correlSheet.LinkToOrigin.LinkSource.EntireRow;
                SheetType sourceType = ExtensionMethods.GetSheetType(correlSheet.LinkToOrigin.LinkSource.Worksheet);
                DisplayCoords dc = DisplayCoords.ConstructDisplayCoords(sourceType);
                string parentID = Convert.ToString(parentRow.Cells[1, dc.ID_Offset].value);
                StringBuilder subIDs = new StringBuilder();
                Excel.Range matrixEnd = correlSheet.xlMatrixCell.End[Excel.XlDirection.xlToRight];
                matrixEnd = matrixEnd.End[Excel.XlDirection.xlDown];
                Excel.Range fieldEnd = correlSheet.xlMatrixCell.End[Excel.XlDirection.xlToRight];
                object[,] matrixVals = correlSheet.xlSheet.Range[correlSheet.xlMatrixCell.Offset[1, 0], matrixEnd].Value;
                int numberOfInputs = matrixVals.GetLength(0);

                header.Append(numberOfInputs);
                header.Append(",");
                header.Append("DM");
                header.Append(",");
                header.Append(parentID);

                for(int row = 1; row <= matrixVals.GetLength(0) - 1; row++)
                {
                    values.Append("&");
                    for(int col = row+1; col <= matrixVals.GetLength(1); col++)
                    {
                        values.Append(matrixVals[row, col].ToString());
                        values.Append(",");
                    }
                    values.Remove(values.Length - 1, 1);    //remove the final ","
                }

                this.Value = $"{header.ToString()}{values.ToString()}";
            }


            public CorrelationString_DM(string parent_id, object[] sub_ids, object[] sub_fields, Data.CorrelationMatrix matrix)
            {
                this.Value = ExtensionMethods.CleanStringLinebreaks(CreateValue(parent_id, sub_ids, sub_fields, matrix.GetMatrix_Values()));
            }

            public static bool Validate()
            {
                return true;
            }

            public override double[,] GetMatrix_Doubles()
            {
                string myValue = ExtensionMethods.CleanStringLinebreaks(this.Value);
                string[] lines = DelimitString(myValue);
                string[] header = lines[0].Split(',');
                int length = Convert.ToInt32(header[0]);
                double[,] matrix = new double[length, length];

                for (int row = 0; row < length; row++)
                {
                    string[] values;
                    if (row + 1 < length)
                        values = lines[row + 1].Split(',');       //broken by entry
                    else
                        values = null;

                    for (int col = row; col < length; col++)
                    {
                        if (col == row)
                            matrix[row, col] = 1;
                        else if (col > row && values != null)
                        {
                            if (Double.TryParse(values[(col - row) - 1], out double conversion))
                            {
                                matrix[row, col] = conversion;
                            }
                            else
                            {
                                throw new Exception("Malformed correlation string");
                            }
                        }

                    }
                }

                //Fill in lower triangular
                for (int row = 0; row < length; row++)
                {
                    for (int col = 0; col < row; col++)
                    {
                        matrix[row, col] = matrix[col, row];
                    }
                }

                return matrix;
            }


            public override string[] GetIDs()
            {
                //HEADER: # INPUTS, TYPE, PARENT_ID, SUB_ID1 ... SUB_IDn
                string[] correlLines = DelimitString(this.Value);
                string[] header = correlLines[0].Split(',');            //get fields (first line) and delimit
                string parentID = header[2];
                string[] returnIDs = new string[header.Length - 3];
                for (int i = 3; i < header.Length; i++)
                    returnIDs[i - 3] = header[i];
                return returnIDs;
            }

            protected override string CreateValue(string parentID, object[] ids, object[] fields, object[,] correlArray)
            {
                correlArray = ExtensionMethods.ReIndexArray<object>(correlArray);
                StringBuilder sb = new StringBuilder();
                sb.Append($"{ids.Length},DM,");
                sb.Append(parentID);
                for (int i = 0; i < ids.Length; i++)
                {
                    sb.Append(",");
                    sb.Append(ids[i]);
                }
                sb.AppendLine();
                //for (int field = 0; field < fields.Length; field++)
                //{
                //    //Add fields
                //    if (field > 0)
                //        sb.Append(",");
                //    sb.Append(fields[field]);
                //}
                //sb.AppendLine();
                for (int row = 0; row < correlArray.GetLength(0); row++)
                {
                    for (int col = row + 1; col < correlArray.GetLength(1); col++)
                    {
                        sb.Append(correlArray[row, col]);
                        if (col < correlArray.GetLength(1) - 1)
                            sb.Append(",");
                    }
                    if (row < correlArray.GetLength(0) - 2)
                        sb.AppendLine();
                }
                return sb.ToString();
            }

            //COLLAPSE
            public override void PrintToSheet(Excel.Range[] xlCells)
            {
                //Clean the string
                //Split the string by lines
                //Print it to the xlCells

                this.Value = ExtensionMethods.CleanStringLinebreaks(this.Value);
                List<Excel.Range> xlFragments = xlCells.ToList();
                string[] lines = this.Value.Split('&');
                int min;
                if (lines.Count() <= xlCells.Count())
                    min = lines.Count();
                else
                    min = xlCells.Count();
                for (int i = 0; i < min; i++)
                {
                    xlFragments[i].Value = lines[i];
                    xlFragments[i].NumberFormat = "\"Sch Correl\";;;\"CORREL\"";
                }
                xlFragments[0].EntireColumn.ColumnWidth = 10;
            }
        }
    }
}
