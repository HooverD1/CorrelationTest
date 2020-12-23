using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace CorrelationTest
{
    public class PhasingTriple
    {
        public UniqueID uID { get; set; }
        public double TopLeft { get; set; }
        public double DiagonalMultiplier { get; set; }
        public double VerticalMultiplier { get; set; }
        private Data.CorrelationMatrix CorrelMatrix { get; set; }

        public PhasingTriple(Excel.Range xlUIdCell, Excel.Range tripleRange) : this((string)xlUIdCell.Value, (string)tripleRange.Value) { }
        
        public PhasingTriple(string uidString, string triple)
        {
            this.uID = new UniqueID(uidString);
            object[,] tripleValues = SplitTriple(triple);
            if (!ValidateTriple(tripleValues))
                throw new Exception("Invalid phasing correlation triple.");
            else
            {
                this.TopLeft = Convert.ToDouble(tripleValues[0,0]);
                this.DiagonalMultiplier = Convert.ToDouble(tripleValues[0,1]);
                this.VerticalMultiplier = Convert.ToDouble(tripleValues[0,2]);
            }
        }

        private string[,] SplitTriple(string triple)
        {
            string[] splitValues = triple.Split(',');
            string[,] tripleValues = new string[1, splitValues.Length];
            for(int i = 0; i < splitValues.Length; i++)
            {
                tripleValues[0, i] = splitValues[i];
            }
            return tripleValues;
        }

        private bool ValidateTriple(object[,] tripleValues)
        {
            //Make sure the values make sense
            //Check shape
            if (tripleValues.GetLength(0) != 1)
                return false;
            if (tripleValues.GetLength(1) != 3)
                return false;
            
            for(int i = 0; i < tripleValues.GetLength(1); i++)
            {
                double tempResult;
                if (!double.TryParse(tripleValues[0, i].ToString(), out tempResult))        //Make sure they convert to doubles
                    return false;
                if (i == 0)
                {
                    if (tempResult > 1 || tempResult < -1)
                        return false;
                }
                else
                {
                    if (tempResult < 0 || tempResult > 1)
                        return false;
                }
            }
            return true;
        }

        public override string ToString()
        {
            return $"{TopLeft},{DiagonalMultiplier},{VerticalMultiplier}";
        }

        public Data.CorrelationMatrix GetPhasingCorrelationMatrix(int periods)
        {
            if (CorrelMatrix == null)
            {
                object[,] matrix = new object[periods, periods];
                for (int row = 0; row < periods; row++)
                {
                    for (int col = row; col < periods; col++)
                    {
                        if (row == col)
                            matrix[row, col] = 1;
                        else
                            matrix[row, col] = TopLeft * Math.Pow(DiagonalMultiplier, col - 1) * Math.Pow(VerticalMultiplier, col - row - 1);
                    }
                }
                this.CorrelMatrix = new Data.CorrelationMatrix(this.uID, matrix);
            }
            return CorrelMatrix;
        }

        public void PrintToCell(Excel.Range xlCell)
        {
            xlCell.Value = $"{TopLeft}, {DiagonalMultiplier}, {VerticalMultiplier}";
        }
    }
}
