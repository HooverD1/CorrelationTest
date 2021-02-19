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
        public class CorrelationString_DT : CorrelationString
        {
            public Triple InputTriple { get; set; }

            public CorrelationString_DT(string correlStringValue)
            {
                this.Value = ExtensionMethods.CleanStringLinebreaks(correlStringValue);
                string triple = this.Value.Split('&')[1];
                this.InputTriple = new Triple(this.GetParentID().ID, triple);
            }

            public CorrelationString_DT(Sheets.CorrelationSheet_Duration correlSheet)
            {
                StringBuilder header = new StringBuilder();
                StringBuilder fields = new StringBuilder();
                StringBuilder values = new StringBuilder();

                Excel.Range parentRow = correlSheet.LinkToOrigin.LinkSource.EntireRow;
                SheetType sourceType = ExtensionMethods.GetSheetType(correlSheet.LinkToOrigin.LinkSource.Worksheet);
                DisplayCoords dc = DisplayCoords.ConstructDisplayCoords(sourceType);
                string parentID = Convert.ToString(parentRow.Cells[1, dc.ID_Offset].value);
                string tripleString = Convert.ToString(correlSheet.xlTripleCell.Value);
                Triple triple = new Triple(tripleString);
                StringBuilder subIDs = new StringBuilder();
                Excel.Range matrixEnd = correlSheet.xlMatrixCell.End[Excel.XlDirection.xlToRight];
                matrixEnd = matrixEnd.End[Excel.XlDirection.xlDown];
                Excel.Range fieldEnd = correlSheet.xlMatrixCell.End[Excel.XlDirection.xlToRight];
                object[,] matrixVals = correlSheet.xlSheet.Range[correlSheet.xlMatrixCell.Offset[1, 0], matrixEnd].Value;
                object[,] fieldVals2D = correlSheet.xlSheet.Range[correlSheet.xlMatrixCell, fieldEnd].Value;
                fieldVals2D = ExtensionMethods.ReIndexArray(fieldVals2D);
                object[] fieldVals = ExtensionMethods.ToJaggedArray(fieldVals2D)[0];
                int numberOfInputs = matrixVals.GetLength(0);

                header.Append(numberOfInputs);
                header.Append(",");
                header.Append("DT");
                header.Append(",");
                header.Append(parentID);

                //foreach (object field in fieldVals)
                //{
                //    fields.Append(Convert.ToString(field));
                //    fields.Append(",");
                //}
                //fields.Remove(fields.Length - 1, 1);    //remove the final char

                values.Append(triple.GetValuesString());

                //This code to convert to matrix:
                /*
                for (int row = 1; row < matrixVals.GetLength(0) - 1; row++)
                {
                    for (int col = row + 1; col < matrixVals.GetLength(1); col++)
                    {
                        values.Append(matrixVals[row, col]);
                        values.Append(",");
                    }
                    values.Remove(values.Length - 1, 1);    //remove the final char
                }
                */
                this.Value = $"{header.ToString()}&{values.ToString()}";
            }


            public CorrelationString_DT(string[] fields, Triple it, string parent_id, string[] sub_ids)        //build a triple string out of a triple
            {
                this.InputTriple = it;
                StringBuilder sb = new StringBuilder();
                sb.Append($"{fields.Length},DT,{parent_id}");
                for (int j = 0; j < sub_ids.Length; j++)
                {
                    sb.Append(",");
                    sb.Append(sub_ids[j]);
                }
                sb.AppendLine();
                //for (int i = 0; i < fields.Length - 1; i++)
                //{
                //    sb.Append(fields[i]);
                //    sb.Append(",");
                //}
                //sb.Append(fields[fields.Length - 1]);
                //sb.AppendLine();
                sb.Append(it.ToString());
                this.Value = ExtensionMethods.CleanStringLinebreaks(sb.ToString());
            }

            public static bool Validate()
            {
                return true;
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

            public override void Expand(Excel.Range xlSource)
            {
                //construct the correlSheet
                Sheets.CorrelationSheet correlSheet = Sheets.CorrelationSheet.Construct(this, xlSource, new Data.CorrelSheetSpecs(SheetType.Correlation_DT));
                //print the correlSheet                         //CorrelationSheet NEEDS NEW CONSTRUCTORS BUILT FOR NON-INPUTS
                correlSheet.PrintToSheet();
            }

            public override object[,] GetMatrix(string[] fields)
            {
                return this.InputTriple.GetCorrelationMatrix(this.GetParentID().ID, this.GetIDs(), fields, SheetType.Correlation_DT).Matrix;
            }


            public Triple GetTriple()
            {
                string[] correlLines = DelimitString(this.Value);
                if (correlLines.Length != 2)
                    throw new Exception("Malformed triple string.");
                string uidString = correlLines[0].Split(',')[2];
                string tripleString = correlLines[1];
                return new Triple(uidString, tripleString);
            }

            public override UniqueID GetParentID()
            {
                string[] lines = this.Value.Split('&');
                return UniqueID.ConstructFromExisting(lines[0]);
            }

            public override void PrintToSheet(Excel.Range[] xlFragments)
            {
                //Clean the string
                //Split the string by lines
                //Print it to the xlCells

                this.Value = ExtensionMethods.CleanStringLinebreaks(this.Value);
                
                string[] lines = this.Value.Split('&');
                int min;
                if (lines.Count() <= xlFragments.Count())
                    min = lines.Count();
                else
                    min = xlFragments.Count();
                for (int i = 0; i < min; i++)
                {
                    xlFragments[i].Value = lines[i];
                    xlFragments[i].NumberFormat = "\"Sch Correl\";;;\"SCH_CORREL\"";
                }
            }
        }
    }
    
}
