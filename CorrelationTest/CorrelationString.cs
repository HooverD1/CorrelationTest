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
        public class CorrelationString
        {
            public string Value { get; set; }

            public void PrintToSheet(Excel.Range xlCell)
            {
                xlCell.Value = this.Value;
                xlCell.NumberFormat = "\"Correl\";;;\"CORREL\"";
            }

            public virtual object[] GetFields() { return null; }
            public virtual UniqueID[] GetIDs() { return null; }
            

            protected CorrelationString() { }

            public CorrelationString(string[] fields)     //creates zero string
            {
                this.Value = CreateValue_Zero(fields);
            }



            public string CreateValue_Zero(string[] fields, double defaultValue = 0) //create a zero correlstring from very generic params
            {
                StringBuilder sb = new StringBuilder();
                for (int i = 0; i < fields.Length; i++)
                {
                    sb.Append(fields[i]);
                    if (i < fields.Length - 1)
                        sb.Append(",");
                }
                sb.AppendLine();
                for (int row = 0; row < fields.Length-1; row++)
                {
                    for (int col = row+1; col < fields.Length; col++)
                    {
                        if(col > row)
                        {
                            sb.Append(defaultValue.ToString());
                        }
                        else
                        {
                            continue;
                        }
                        if (col < fields.Length-1)
                            sb.Append(",");
                    }
                    if (row < fields.Length - 2)
                        sb.AppendLine();
                }
                return sb.ToString();
            }

            protected string[] DelimitString()
            {
                string correlString = this.Value.Replace("\r\n", "&");  //simplify delimiter
                correlString = correlString.Replace("\n", "&");  //simplify delimiter
                string[] correlLines = correlString.Split('&');         //split lines
                return correlLines;
            }

            public object[,] GetMatrix()
            {       //returning 2,2 instead of 3,3
                string myValue = this.Value.Replace("\r\n", "&");
                myValue = myValue.Replace("\n", "&");
                //myValue = myValue + "&";    //add a blank final row in for the sake of the array
                string[] fieldString1 = myValue.Split('&');          //broken by line
                string[] fieldString = new string[fieldString1.Length - 1];
                for(int i = 1; i < fieldString1.Length; i++) { fieldString[i - 1] = fieldString1[i]; }
                object[,] matrix = new object[fieldString.Length+1, fieldString.Length+1];

                for (int row = 0; row < fieldString.Length+1; row++)
                {
                    string[] values;
                    if (row < fieldString.Length)
                        values = fieldString[row].Split(',');       //broken by entry
                    else
                        values = null;

                    for (int col = fieldString.Length; col >= 0; col--)
                    {
                        if (col == row)
                            matrix[row, col] = 1;
                        else if (col > row)
                            matrix[row, col] = Convert.ToDouble(values[(col - row) - 1]);

                        else  //col < row
                            matrix[row, col] = null;
                    }
                }
                return matrix;
            }

            protected virtual string CreateValue(UniqueID[] ids, object[,] correlArray)
            {
                correlArray = ExtensionMethods.ReIndexArray<object>(correlArray);
                StringBuilder sb = new StringBuilder();

                for (int field = 0; field < correlArray.GetLength(1); field++)
                {
                    //Add fields
                    sb.Append(ids[field].ID);
                    if (field < correlArray.GetLength(1) - 1)
                        sb.Append(",");
                }
                sb.AppendLine();
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
        }
    }
    
}
