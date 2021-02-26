﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace CorrelationTest
{
    namespace Data
    {
        public class CorrelationString_PP : CorrelationString
        {
            public Triple Triple { get; set; }
            public CorrelationString_PP(Triple pt, string[] start_dates, string parent_id)        //build a triple string out of a triple
            {
                this.Triple = pt;
                StringBuilder sb = new StringBuilder();
                sb.Append($"{start_dates.Length},PT,{parent_id}");  //Header
                sb.AppendLine();
                //foreach (string start_date in start_dates)
                //{
                //    sb.Append(start_date);    //Period start dates as fields
                //    sb.Append(",");
                //}
                //sb.Remove(sb.Length - 1, 1);    //remove the final comma on the fields
                //sb.AppendLine();
                sb.Append(pt.ToString());
                this.Value = ExtensionMethods.CleanStringLinebreaks(sb.ToString());
            }

            public CorrelationString_PP(Sheets.CorrelationSheet_Phasing correlSheet)
            {
                StringBuilder header = new StringBuilder();
                StringBuilder fields = new StringBuilder();
                StringBuilder values = new StringBuilder();

                Excel.Range parentRow = correlSheet.LinkToOrigin.LinkSource.EntireRow;
                SheetType sourceType = ExtensionMethods.GetSheetType(correlSheet.LinkToOrigin.LinkSource.Worksheet);
                DisplayCoords dc = DisplayCoords.ConstructDisplayCoords(sourceType);
                string parentID = Convert.ToString(parentRow.Cells[1, dc.ID_Offset].value);
                string tripleString = Convert.ToString(correlSheet.xlPairsCell.Value);
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
                header.Append("PT");
                header.Append(",");
                header.Append(parentID);

                foreach (object field in fieldVals)
                {
                    fields.Append(Convert.ToString(field));
                    fields.Append(",");
                }
                fields.Remove(fields.Length - 1, 1);    //remove the final char

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


            public CorrelationString_PP(string correlString)
            {
                this.Value = ExtensionMethods.CleanStringLinebreaks(correlString);
                string triple = this.Value.Split('&')[1];
                this.Triple = new Triple(this.GetParentID().ID, triple);
            }

            public override object[,] GetMatrix(string[] fields)
            {
                return this.Triple.GetCorrelationMatrix(fields);
            }

            public override void PrintToSheet(Excel.Range xlPhasingCorrelCell)
            {
                //Clean the string
                //Split the string by lines
                //Print it to the xlCells

                this.Value = ExtensionMethods.CleanStringLinebreaks(this.Value);
                xlPhasingCorrelCell.Value = this.Value;
                xlPhasingCorrelCell.NumberFormat = "\"Ph Correl\";;;\"PH_CORREL\"";
                //string[] lines = this.Value.Split('&');
                //int min;
                //if (lines.Count() <= xlFragments.Count())
                //    min = lines.Count();
                //else
                //    min = xlFragments.Count();
                //int lineIndex = 0;
                //for (int j = 0; j < xlFragments.Count(); j++)
                //{   //Iterate Areas
                //    for (int i = 0; i < xlFragments[j].Cells.Count; i++)
                //    {   //Iterate cells within areas
                //        if (lineIndex < lines.Count())
                //        {
                //            xlFragments[j].Cells[i, 1].Value = lines[lineIndex++];
                //            xlFragments[j].Cells[i, 1].NumberFormat = "\"Ph Correl\";;;\"PH_CORREL\"";
                //        }
                //        else
                //        {
                //            //Remaining cells stay empty
                //            break;
                //        }
                //    }
                //}
            }

            public override UniqueID GetParentID()
            {
                string[] lines = CorrelationString.DelimitString(this.Value);
                string[] header = lines[0].Split(',');
                return UniqueID.ConstructFromExisting(header[2]);
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

            public override string[] GetIDs()
            {
                var period_ids = PeriodID.GeneratePeriodIDs(this.GetParentID(), this.GetNumberOfSubs());
                return period_ids.Select(x => x.ID).ToArray();
            }

            public static bool Validate()
            {
                return true;
            }
        }
    }    
}