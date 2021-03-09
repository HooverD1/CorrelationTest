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
            private int matrixRow { get; }
            private int matrixCol { get; }
            private int linkRow { get; }
            private int linkCol { get; }
            private string linkFormat { get; }
            private int idRow { get; }
            private int idCol { get; }
            private int subIdRow { get; }
            private int subIdCol { get; }
            private int distRow { get; }
            private int distCol { get; }
            private int stringRow { get; }
            private int stringCol { get; }
            private int pairsRow { get; }
            private int pairsCol { get; }

            public Tuple<int, int> MatrixCoords { get; }
            public Tuple<int, int> MatrixCoords_End { get; set; }
            public Tuple<int, int> LinkCoords { get; }
            public Tuple<int, int> IdCoords { get; }
            public Tuple<int, int> SubIdCoords { get; }
            public Tuple<int, int> DistributionCoords { get; }
            public Tuple<int, int> StringCoords { get; }
            public Tuple<int, int> PairsCoords { get; }
            public string LinkFormat { get; }

            public CorrelSheetSpecs(SheetType correlSheetType)
            {
                switch (correlSheetType)
                {
                    case SheetType.Correlation_CM:
                        matrixRow = 4;
                        matrixCol = 4;
                        linkRow = 3;
                        linkCol = 1;
                        linkFormat = "\"Correl\";;;\"CORREL\"";
                        idRow = 4;
                        idCol = 1;
                        subIdRow = 5;
                        subIdCol = 2;
                        distRow = 5;
                        distCol = 1;
                        stringRow = 2;
                        stringCol = 1;

                        break;
                    case SheetType.Correlation_CP:
                        matrixRow = 4;
                        matrixCol = 9;
                        linkRow = 3;
                        linkCol = 1;
                        linkFormat = "\"Correl\";;;\"CORREL\"";
                        idRow = 4;
                        idCol = 1;
                        subIdRow = 5;
                        subIdCol = 5;
                        distRow = 5;
                        distCol = 1;
                        stringRow = 2;
                        stringCol = 1;
                        pairsRow = 5;
                        pairsCol = 6;       //Requires a 2 cell width
                        break;
                    case SheetType.Correlation_PM:
                        matrixRow = 4;
                        matrixCol = 4;
                        linkRow = 3;
                        linkCol = 1;
                        linkFormat = "\"Correl\";;;\"CORREL\"";
                        idRow = 4;
                        idCol = 1;
                        subIdRow = 5;
                        subIdCol = 2;
                        distRow = 5;
                        distCol = 1;
                        stringRow = 2;
                        stringCol = 1;
                        pairsRow = 2;
                        pairsCol = 2;
                        break;
                    case SheetType.Correlation_PP:
                        matrixRow = 4;
                        matrixCol = 4;
                        linkRow = 3;
                        linkCol = 1;
                        linkFormat = "\"Correl\";;;\"CORREL\"";
                        idRow = 4;
                        idCol = 1;
                        subIdRow = 5;
                        subIdCol = 2;
                        distRow = 5;
                        distCol = 1;
                        stringRow = 2;
                        stringCol = 1;
                        pairsRow = 2;
                        pairsCol = 2;
                        break;
                    case SheetType.Correlation_DM:
                        matrixRow = 4;
                        matrixCol = 4;
                        linkRow = 3;
                        linkCol = 1;
                        linkFormat = "\"Correl\";;;\"CORREL\"";
                        idRow = 4;
                        idCol = 1;
                        subIdRow = 5;
                        subIdCol = 2;
                        distRow = 5;
                        distCol = 1;
                        stringRow = 2;
                        stringCol = 1;
                        pairsRow = 2;
                        pairsCol = 2;
                        break;
                    case SheetType.Correlation_DP:
                        matrixRow = 4;
                        matrixCol = 9;
                        linkRow = 3;
                        linkCol = 1;
                        linkFormat = "\"Correl\";;;\"CORREL\"";
                        idRow = 4;
                        idCol = 1;
                        subIdRow = 5;
                        subIdCol = 5;
                        distRow = 5;
                        distCol = 1;
                        stringRow = 2;
                        stringCol = 1;
                        pairsRow = 5;
                        pairsCol = 6;       //Requires a 2 cell width
                        break;
                    default:
                        throw new Exception("Unknown correl sheet type");
                }
                this.MatrixCoords = new Tuple<int, int>(matrixRow, matrixCol);
                this.LinkCoords = new Tuple<int, int>(linkRow, linkCol);
                this.LinkFormat = linkFormat;
                this.IdCoords = new Tuple<int, int>(idRow, idCol);
                this.SubIdCoords = new Tuple<int, int>(subIdRow, subIdCol);
                this.DistributionCoords = new Tuple<int, int>(distRow, distCol);
                this.StringCoords = new Tuple<int, int>(stringRow, stringCol);
                this.PairsCoords = new Tuple<int, int>(pairsRow, pairsCol);
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
