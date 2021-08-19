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
        public class CorrelationString_CP : CorrelationString
        {
            public PairSpecification Pairs { get; set; }
            public CorrelationString_CP(string correlString)
            {
                if (Validate(correlString))
                {
                    this.Value = ExtensionMethods.CleanStringLinebreaks(correlString);
                    string[] lines = this.Value.Split('&');
                    string triple = lines[1];
                    this.Pairs = PairSpecification.ConstructFromString(correlString);
                }
                else
                    throw new FormatException("Malformed CP string");
            }

            //COLLAPSE
            public CorrelationString_CP(Sheets.CorrelationSheet_CP correlSheet)
            {
                StringBuilder header = new StringBuilder();
                StringBuilder fields = new StringBuilder();
                StringBuilder values = new StringBuilder();


                Excel.Range linkedRange = correlSheet.LinkToOrigin.LinkSource;
                Excel.Range parentRow = linkedRange.EntireRow;
                SheetType sourceType = ExtensionMethods.GetSheetType(linkedRange.Worksheet);
                DisplayCoords dc = DisplayCoords.ConstructDisplayCoords(sourceType);
                string parentID = Convert.ToString(parentRow.Cells[1, dc.ID_Offset].value);
                int size = Data.CorrelationString.GetNumberOfInputsFromCorrelStringValue(correlSheet.xlHeaderCell.Value);
                PairSpecification pairs = PairSpecification.ConstructFromRange(correlSheet.xlPairsCell, size - 1);
                Excel.Range matrixEnd = correlSheet.xlMatrixCell.End[Excel.XlDirection.xlToRight];
                matrixEnd = matrixEnd.End[Excel.XlDirection.xlDown];
                Excel.Range fieldEnd = correlSheet.xlMatrixCell.End[Excel.XlDirection.xlToRight];
                object[,] matrixVals = correlSheet.xlSheet.Range[correlSheet.xlMatrixCell.Offset[1, 0], matrixEnd].Value;
                object[,] fieldVals2D = correlSheet.xlSheet.Range[correlSheet.xlMatrixCell, fieldEnd].Value;
                fieldVals2D = ExtensionMethods.ReIndexArray(fieldVals2D);
                object[] fieldVals = ExtensionMethods.ToJaggedArray(fieldVals2D)[0];
                int numberOfInputs = matrixVals.GetLength(0);

                CostSheet costSheet = CostSheet.ConstructFromXlCostSheet(linkedRange.Worksheet);
                IHasCostCorrelations parentItem = (IHasCostCorrelations)(from Item parent in costSheet.Items where parent.uID.ID == parentID select parent).First();
                IEnumerable<string> subStrings = from ISub sub in parentItem.SubEstimates select sub.uID.ID;

                header.Append(numberOfInputs);
                header.Append(",");
                header.Append("CP");
                header.Append(",");
                header.Append(parentID);

                foreach(string subString in subStrings)
                {
                    header.Append(",");
                    header.Append(subString);
                }

                foreach (object field in fieldVals)
                {
                    fields.Append(Convert.ToString(field));
                    fields.Append(",");
                }
                fields.Remove(fields.Length - 1, 1);    //remove the final char

                values.Append(pairs.Value);

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


            public CorrelationString_CP(string[] fields, PairSpecification ps, string parent_id, string[] sub_ids)        //build a triple string out of a triple
            {
                this.Pairs = ps;
                StringBuilder sb = new StringBuilder();
                sb.Append($"{fields.Length},CP,{parent_id}");
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
                //sb.Append(fields[fields.Length-1]);
                //sb.AppendLine();
                sb.Append(ps.ToString());
                this.Value = ExtensionMethods.CleanStringLinebreaks(sb.ToString());
            }

            //public PairSpecification GetPairs()
            //{
            //    return PairSpecification.ConstructFromString(this.Value);
            //}

            public PairSpecification GetPairwise()
            {
                string[] correlLines = DelimitString(this.Value);
                string uidString = correlLines[0].Split(',')[2];
                return PairSpecification.ConstructFromString(this.Value);
            }

            public override string[,] GetMatrix_Formulas(Sheets.CorrelationSheet CorrelSheet)
            {
                return this.Pairs.GetCorrelationMatrix_Formulas(CorrelSheet);
            }

            public override double[,] GetMatrix_Doubles()
            {
                return this.Pairs.GetCorrelationMatrix_Values();
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

            public static bool Validate()
            {
                return true;
            }

            protected override bool Validate(string correlString)
            {
                string[] lines = correlString.Split('&');
                //Should have at least a header row and a data row

                if (lines.Length != 2)
                    return false;

                //VALIDATE HEADER
                string[] header = lines[0].Split(',');
                //Header should have at least # inputs, type, parent, and two subs (5 items)
                if (header.Length < 5)      //Minimum length
                    return false;
                if (int.TryParse(header[0], out int inputs))        //Check if the lengths are consistent with the "# inputs"
                {   //Expect at least two sub IDs
                    if (inputs < 2)
                        return false;
                    else if (inputs != lines.Length) // 5 subs = header + 4 lines for sub correlation
                        return false;
                    else if (header.Length != inputs + 3)    //Check the header length is consistent
                        return false;
                }
                else
                {   //Can't parse '# inputs' position as an integer
                    return false;
                }

                //VALIDATE THE TYPE
                if (header[1] != "CP")
                {
                    return false;
                }

                //VALIDATE THE IDs
                for (int id = 2; id < header.Length; id++)
                {
                    if (!UniqueID.Validate(header[id]))
                        return false;
                }

                //VALIDATE DATA ROWS AS DOUBLES
                for (int line = 1; line < lines.Length; line++)
                {
                    string[] data_positions = lines[line].Split(',');
                    if (data_positions.Length != 2)
                        return false;

                    if (!Double.TryParse(data_positions[0], out double pos0Value))
                        return false;
                    if (pos0Value > 1 || pos0Value < -1)
                        return false;

                    if (!Double.TryParse(data_positions[1], out double pos1Value))
                        return false;
                    if (pos1Value > 1 || pos1Value < -1)
                        return false;
                }
                return true;
            }

            public override UniqueID GetParentID()
            {            
                string[] lines = this.Value.Split('&');
                return UniqueID.ConstructFromExisting(lines[0]);
            }

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
                    xlFragments[i].NumberFormat = "\"In Correl\";;;\"CORREL\"";
                }
                xlFragments[0].EntireColumn.ColumnWidth = 10;
            }

        }
    }
}
