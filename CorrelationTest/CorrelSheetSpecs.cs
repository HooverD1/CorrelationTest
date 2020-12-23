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
        public class CorrelSheetSpecs
        {
            public Tuple<int, int> MatrixCoords { get; }
            public Tuple<int, int> MatrixCoords_End { get; set; }
            public Tuple<int, int> LinkCoords { get; }
            public Tuple<int, int> IdCoords { get; }
            public Tuple<int, int> DistributionCoords { get; }
            public Tuple<int, int> CorrelTypeCoords { get; }
            public Tuple<int, int> CorrelStringCoords { get; }

            public string LinkFormat { get; }

            public CorrelSheetSpecs(int matrixRow = 4, int matrixCol = 5, int linkRow = 3, int linkCol = 1, string linkFormat = "\"Correl\";;;\"CORREL\"", int idRow = 4, int idCol = 1, int distRow = 5, int distCol = 1, int typeRow = 4, int typeCol = 2, int stringRow = 3, int stringCol = 2)
            {
                this.MatrixCoords = new Tuple<int, int>(matrixRow, matrixCol);
                this.LinkCoords = new Tuple<int, int>(linkRow, linkCol);
                this.LinkFormat = linkFormat;
                this.IdCoords = new Tuple<int, int>(idRow, idCol);
                this.DistributionCoords = new Tuple<int, int>(distRow, distCol);
                this.CorrelTypeCoords = new Tuple<int, int>(typeRow, typeCol);
                this.CorrelStringCoords = new Tuple<int, int>(stringRow, stringCol);
            }

            public void PrintLinkCoords(Excel.Worksheet xlSheet)
            {
                xlSheet.Cells[Sheets.CorrelationSheet.param_RowLink.Item1, Sheets.CorrelationSheet.param_RowLink.Item2] = LinkCoords.Item1;
                xlSheet.Cells[Sheets.CorrelationSheet.param_ColLink.Item1, Sheets.CorrelationSheet.param_ColLink.Item2] = LinkCoords.Item2;
            }

            public void PrintMatrixCoords(Excel.Worksheet xlSheet)
            {
                xlSheet.Cells[Sheets.CorrelationSheet.param_RowMatrix.Item1, Sheets.CorrelationSheet.param_RowMatrix.Item2] = MatrixCoords.Item1;
                xlSheet.Cells[Sheets.CorrelationSheet.param_ColMatrix.Item1, Sheets.CorrelationSheet.param_ColMatrix.Item2] = MatrixCoords.Item2;
            }

            public void PrintIdCoords(Excel.Worksheet xlSheet)
            {
                xlSheet.Cells[Sheets.CorrelationSheet.param_RowID.Item1, Sheets.CorrelationSheet.param_RowID.Item2] = IdCoords.Item1;
                xlSheet.Cells[Sheets.CorrelationSheet.param_ColID.Item1, Sheets.CorrelationSheet.param_ColID.Item2] = IdCoords.Item2;
            }

            public void PrintDistCoords(Excel.Worksheet xlSheet)
            {
                xlSheet.Cells[Sheets.CorrelationSheet.param_RowDist.Item1, Sheets.CorrelationSheet.param_RowDist.Item2] = DistributionCoords.Item1;
                xlSheet.Cells[Sheets.CorrelationSheet.param_ColDist.Item1, Sheets.CorrelationSheet.param_ColDist.Item2] = DistributionCoords.Item2;
            }
        }
    }
}
