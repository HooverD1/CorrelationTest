using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using Vsto = Microsoft.Office.Tools.Excel;

namespace CorrelationTest
{
    namespace Sheets
    {
        public class CorrelationSheet_PP : CorrelationSheet
        {
            public Data.CorrelationString_PP CorrelString { get; set; }
            public PairSpecification PairSpec { get; set; }

            //EXPAND
            public CorrelationSheet_PP(IHasPhasingCorrelations ParentItem)        //bring in the coordinates and set up the ranges once they exist
            {   //Build from the correlString to get the xlSheet
                this.CorrelString = (Data.CorrelationString_PP)ParentItem.PhasingCorrelationString;
                this.Specs = new Data.CorrelSheetSpecs(SheetType.Correlation_PP);
                this.xlSheet = GetXlSheet();
                this.LinkToOrigin = new Data.Link(ParentItem.xlCorrelCell_Phasing);
                this.xlLinkCell = xlSheet.Cells[Specs.LinkCoords.Item1, Specs.LinkCoords.Item2];
                this.xlHeaderCell = xlSheet.Cells[Specs.StringCoords.Item1, Specs.StringCoords.Item2];
                //this.xlIDCell = xlSheet.Cells[Specs.IdCoords.Item1, Specs.IdCoords.Item2];
                this.xlSubIdCell = xlSheet.Cells[Specs.SubIdCoords.Item1, Specs.SubIdCoords.Item2];
                this.xlDistCell = xlSheet.Cells[Specs.DistributionCoords.Item1, Specs.DistributionCoords.Item2];
                this.xlMatrixCell = xlSheet.Cells[Specs.MatrixCoords.Item1, Specs.MatrixCoords.Item2];
                this.xlPairsCell = xlSheet.Cells[Specs.PairsCoords.Item1, Specs.PairsCoords.Item2];
                this.xlSubIdCell = xlSheet.Cells[Specs.SubIdCoords.Item1, Specs.SubIdCoords.Item2];
                this.xlButton_CollapseCorrel = xlSheet.Cells[Specs.Btn_Collapse.Item1, Specs.Btn_Collapse.Item2];
                this.xlButton_Cancel = xlSheet.Cells[Specs.Btn_Cancel.Item1, Specs.Btn_Cancel.Item2];
                this.Specs.PrintMatrixCoords(xlSheet);                                          //Print the matrix start coords
                this.Specs.PrintLinkCoords(xlSheet);                                            //Print the link coords
                this.Specs.PrintIdCoords(xlSheet);                                              //Print the ID coords
                this.Specs.PrintDistCoords(xlSheet);                                            //Print the Distribution coords
                CorrelMatrix = Data.CorrelationMatrix.ConstructFromParentItem(ParentItem, SheetType.Correlation_PP, this);
                this.PrintMatrixEndCoords(xlSheet);                                             //Print the matrix end coords
                this.PairSpec = ((Data.CorrelationString_PP)CorrelString).GetPairwise();
            }

            //COLLAPSE
            public CorrelationSheet_PP() //build from the xlsheet to get the string
            {
                this.xlSheet = GetXlSheet();
                this.Specs = new Data.CorrelSheetSpecs(SheetType.Correlation_PP);
                //Set up the xlCells
                this.xlLinkCell = xlSheet.Cells[Specs.LinkCoords.Item1, Specs.LinkCoords.Item2];
                this.xlHeaderCell = xlSheet.Cells[Specs.StringCoords.Item1, Specs.StringCoords.Item2];
                //this.xlIDCell = xlSheet.Cells[Specs.IdCoords.Item1, Specs.IdCoords.Item2];
                this.xlDistCell = xlSheet.Cells[Specs.DistributionCoords.Item1, Specs.IdCoords.Item2];
                this.xlMatrixCell = xlSheet.Cells[Specs.MatrixCoords.Item1, Specs.MatrixCoords.Item2];
                this.xlPairsCell = xlSheet.Cells[Specs.PairsCoords.Item1, Specs.PairsCoords.Item2];
                this.xlSubIdCell = xlSheet.Cells[Specs.SubIdCoords.Item1, Specs.SubIdCoords.Item2];

                //Set up the link
                this.LinkToOrigin = new Data.Link(Convert.ToString(xlLinkCell.Value));

                //Build the CorrelMatrix
                int fields = Convert.ToInt32(Convert.ToString(xlHeaderCell.Value).Split(',')[0]);
                Excel.Range fieldRange = xlMatrixCell.Resize[1, fields];
                Excel.Range matrixRange = xlMatrixCell.Offset[1, 0].Resize[fields, fields];
                //this.CorrelMatrix = new Data.CorrelationMatrix(this, fieldRange, matrixRange);
                this.CorrelMatrix = Data.CorrelationMatrix.ConstructFromCorrelationSheet(this);
                //Build the CorrelString, which can print itself during collapse
                SheetType sheetType = ExtensionMethods.GetSheetType(xlSheet);
                //Build the triple from the string
                string correlStringVal = this.xlHeaderCell.Value;
                Data.CorrelationString_PP existing_cst = new Data.CorrelationString_PP(correlStringVal);
                this.PairSpec = existing_cst.GetPairwise();
            }

            //public override void UpdateCorrelationString(string[] ids)
            //{
            //    this.CorrelString = new Data.CorrelationString_PM(ids, this.CorrelMatrix);
            //}

            protected override Excel.Worksheet GetXlSheet(bool CreateNew = true)
            {
                var xlCorrelSheets = from Excel.Worksheet sheet in ThisAddIn.MyApp.Worksheets
                                     where sheet.Cells[1, 1].value == "$CORRELATION_PP"
                                     select sheet;
                if (xlCorrelSheets.Any())
                    xlSheet = xlCorrelSheets.First();
                else if (CreateNew)
                    xlSheet = CreateXLCorrelSheet("_PP");
                else
                    throw new Exception("No input matrix correlation sheet found.");
                return xlSheet;
            }

            protected override Excel.Worksheet CreateXLCorrelSheet(string postfix)
            {
                Excel.Worksheet xlCorrelSheet = ThisAddIn.MyApp.Worksheets.Add(After: ThisAddIn.MyApp.ActiveWorkbook.Sheets[ThisAddIn.MyApp.ActiveWorkbook.Sheets.Count]);
                xlCorrelSheet.Name = "Correlation";
                xlCorrelSheet.Cells[1, 1] = $"$CORRELATION{postfix}";
                xlCorrelSheet.Rows[1].Hidden = true;
                return xlCorrelSheet;
            }

            //public override void UpdateCorrelationString(string[] ids)
            //{
            //    UniqueID parentID = UniqueID.ConstructFromExisting(Convert.ToString(this.xlIDCell.Value));
            //    object[,] matrix = this.xlMatrixCell.Offset[1, 0].Resize[ids.Length, ids.Length].Value;
            //    this.CorrelMatrix = Data.CorrelationMatrix.ConstructFromCorrelationSheet(this);
            //    this.CorrelString = new Data.CorrelationString_CM(parentID.ID, ids, this.CorrelMatrix.Fields, CorrelMatrix);
            //    this.xlHeaderCell.Value = this.CorrelString.Value;
            //}

            //Bring in the coordinates - use an enum to build them for each sheet type
            //Parse CorrelString to get type for collapse

            protected override string GetDistributionString(IHasCorrelations est)
            {
                StringBuilder sb = new StringBuilder();
                sb.Append($"{((IHasPhasingCorrelations)est).PhasingDistribution.Name}");
                for (int i = 1; i < ((IHasPhasingCorrelations)est).PhasingDistribution.DistributionParameters.Count(); i++)
                {
                    string param = $"Param{i}";
                    if (((IHasPhasingCorrelations)est).PhasingDistribution.DistributionParameters[param] != null)
                        sb.Append($",{((IHasPhasingCorrelations)est).PhasingDistribution.DistributionParameters[param]}");
                }
                return sb.ToString();
            }

            public override void PrintToSheet()  //expanding from string
            {
                CostSheet costSheet = CostSheet.ConstructFromXlCostSheet(this.LinkToOrigin.LinkSource.Worksheet);
                //Should be some way to pull the already instantiated costSheet in here..
                Item parentItem = (from Item i in costSheet.Items where i.uID.ID == Convert.ToString(this.LinkToOrigin.LinkSource.EntireRow.Cells[1, costSheet.Specs.ID_Offset].value) select i).First();
                IHasPhasingCorrelations parentEstimate;
                if (!(parentItem is IHasPhasingCorrelations) || ((IHasPhasingCorrelations)parentItem).Periods.Count() < 2)
                {
                    throw new Exception("Item has no phasing correlations");
                }
                else
                {
                    parentEstimate = (IHasPhasingCorrelations)parentItem;
                }

                //IHasPhasingCorrelations tempEst = (IHasPhasingCorrelations)Item.ConstructFromRow(this.LinkToOrigin.LinkSource.EntireRow, costSheet);        //Load only this parent estimate
                this.CorrelMatrix.PrintToSheet(xlMatrixCell);                                   //Print the matrix
                this.LinkToOrigin.PrintToSheet(xlLinkCell);                                     //Print the link
                CorrelString.PrintToSheet(xlHeaderCell);
                this.xlDistCell.Value = GetDistributionString(parentEstimate);
                int subCount = parentEstimate.Periods.Count();
                Excel.Range xlPairsRange = xlPairsCell.Resize[subCount - 1, 2];
                xlPairsRange.Value = this.PairSpec.GetValuesString_Split(); //((Data.CorrelationString_PP)CorrelString).GetPairwise().Value;

                FormatSheet();
                AddUserControls();
            }

            public override void FormatSheet()
            {
                int size = this.CorrelMatrix.Fields.Length;
                Excel.Range matrixStart = this.xlMatrixCell.Offset[1, 0];
                Excel.Range matrixRange = matrixStart.Resize[size, size];
                Excel.Range xlPairsRange = this.xlPairsCell.Resize[size - 1, 2];
                Excel.Range diagonal = matrixRange.Cells[1, 1];
                for (int i = 2; i <= size; i++)
                    diagonal = ThisAddIn.MyApp.Union(diagonal, matrixRange.Cells[i, i]);
                foreach (Excel.Range row in matrixRange.Rows)
                {
                    row.Interior.Color = System.Drawing.Color.FromArgb(225, 225, 225);      //Dying right here if you try to do it all at once?
                    row.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    row.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                }

                xlPairsRange.Interior.Color = System.Drawing.Color.FromArgb(255, 255, 190);
                xlPairsRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                xlPairsRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                diagonal.Interior.Color = System.Drawing.Color.FromArgb(0, 0, 0);
                diagonal.Font.Color = System.Drawing.Color.FromArgb(255, 255, 255);
            }

            public override bool Validate() //This needs moved to subclass because the CorrelString implementation was moved to subclass
            {
                bool validateMatrix_to_String = this.CorrelString.ValidateAgainstMatrix(this.CorrelMatrix.Fields);
                //need to get fields from xlSheet fresh, not the object, to validate
                bool validateMatrix_to_xlSheet = this.CorrelMatrix.ValidateAgainstXlSheet(this.Get_xlFields());
                return validateMatrix_to_String && validateMatrix_to_xlSheet;
            }

            private void AddUserControls()
            {
                Vsto.Worksheet vstoSheet = Globals.Factory.GetVstoObject(this.xlSheet);

                //COLLAPSE
                System.Windows.Forms.Button btn_CollapseCorrelation = new System.Windows.Forms.Button();
                btn_CollapseCorrelation.Text = "Save Correlation";
                btn_CollapseCorrelation.Click += CollapseCorrelationClicked;
                vstoSheet.Controls.AddControl(btn_CollapseCorrelation, this.xlButton_CollapseCorrel.Resize[2, 3], "CollapseToCostSheet");

                //CANCEL
                System.Windows.Forms.Button btn_Cancel = new System.Windows.Forms.Button();
                btn_Cancel.Text = "Cancel Changes";
                btn_Cancel.Click += CancelChangesClicked;
                vstoSheet.Controls.AddControl(btn_Cancel, this.xlButton_Cancel.Resize[2, 3], "CancelCorrelationChanges");

            }
        }
    }
}
