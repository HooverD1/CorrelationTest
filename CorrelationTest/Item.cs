using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace CorrelationTest
{
    public abstract class Item
    {
        public Excel.Range xlRow { get; set; }
        public Excel.Range xlTypeCell { get; set; }
        public Excel.Range xlNameCell { get; set; }
        public Excel.Range xlCorrelCell_Cost { get; set; }
        public Excel.Range xlCorrelCell_Duration { get; set; }
        public Excel.Range xlCorrelCell_Phasing { get; set; }
        public Excel.Range xlLevelCell { get; set; }
        public int Level { get; set; }
        public CostSheet ContainingSheetObject { get; set; }
        public string Name { get; set; }
        public UniqueID uID { get; set; }

        protected Item(Excel.Range xlRow, CostSheet ContainingSheetObject)
        {
            this.xlRow = xlRow;
            this.ContainingSheetObject = ContainingSheetObject;
            this.xlTypeCell = xlRow.Cells[1, ContainingSheetObject.Specs.Type_Offset];
            this.xlNameCell = xlRow.Cells[1, ContainingSheetObject.Specs.Name_Offset];
            //this.xlCorrelCell_Cost = xlRow.Cells[1, ContainingSheetObject.Specs.CostCorrel_Offset];
            //this.xlCorrelCell_Phasing = xlRow.Cells[1, ContainingSheetObject.Specs.PhasingCorrel_Offset];
            //this.xlCorrelCell_Duration = xlRow.Cells[1, ContainingSheetObject.Specs.DurationCorrel_Offset];
            LoadUniqueID();
            if(ContainingSheetObject is Sheets.WBSSheet)
            {
                this.xlLevelCell = xlRow.Cells[1, ContainingSheetObject.Specs.Level_Offset];
                if(int.TryParse(Convert.ToString(xlLevelCell.Value),out int level))
                    this.Level = level;
            }
            this.Name = Convert.ToString(xlNameCell.Value);

        }

        protected Item() { }

        protected void LoadUniqueID()
        {
            this.uID = GetUniqueID();
        }

        private UniqueID GetUniqueID()
        {
            object idCellValue = xlRow.Cells[1, ContainingSheetObject.Specs.ID_Offset].value;
            if (idCellValue == null)
            {
                if (this is Input_Item)
                    return UniqueID.ConstructNew("E");
                else if (this is Estimate_Item && ContainingSheetObject is Sheets.EstimateSheet)
                    return UniqueID.ConstructNew("E");
                else if (this is Estimate_Item && ContainingSheetObject is Sheets.WBSSheet)
                    return UniqueID.ConstructNew("W");
                else if (this is Sum_Item)
                    return UniqueID.ConstructNew("S");
                else if (this is WBS_Item)
                    return UniqueID.ConstructNew("W");
                else
                    throw new Exception("Unknown sheet origin type");
            }
            else
            {
                return UniqueID.ConstructFromExisting(Convert.ToString(idCellValue));
            }
        }

        public virtual bool CanExpand(CorrelationType correlType)
        {
            //!!!!!!!!!!! 
            //This needs refactored as overridden by subclasses to avoid all this switching and casting code
            //!!!!!!!!!!!

            /*
             * This method checks whether the given item has anything to expand into a correlation sheet
             */

            if (correlType == CorrelationType.Cost || correlType == CorrelationType.Duration)
            {
                if (this is IHasSubs)
                {
                    if (((IHasSubs)this).SubEstimates.Count <= 1)
                    {
                        return false;
                    }
                    if (this is ISub)
                    {
                        if (((ISub)this).Parent is IJointEstimate)
                        {
                            ExtensionMethods.TurnOnUpdating();
                            return false;
                        }
                    }
                    //Invalid selection
                    //Don't throw an error, just don't do anything.
                    return true;
                }
                else
                {
                    ExtensionMethods.TurnOnUpdating();
                    return false;
                }
            }
            else if (correlType == CorrelationType.Phasing)
            {
                if (this is Input_Item)
                {
                    return false;
                }
                else
                {
                    return true;
                }
            }
            else
            {
                return false;
            }
        }

        public static Item ConstructFromRow(Excel.Range xlRow, CostSheet containing_sheet_object)
        {
            //What kind of sheet object is constructing it?
            //What kind of CostRow is this?
            int type_offset = containing_sheet_object.Specs.Type_Offset;
            string type_value = Convert.ToString(xlRow.Cells[1, type_offset].value);
            CostItems costRow_type = CostItems.Null;
            foreach(CostItems ci in Enum.GetValues(typeof(CostItems)))      //Get the cell value into an enum
            {
                if (type_value == ci.ToString())
                {
                    costRow_type = ci;
                    break;
                }                    
            }
            switch (costRow_type)       //Construct subtypes based on the enum
            {
                case CostItems.CE:
                    return new CostEstimate(xlRow, containing_sheet_object);
                case CostItems.SE:
                    return new ScheduleEstimate(xlRow, containing_sheet_object);
                case CostItems.CASE:
                    return new CostScheduleEstimate(xlRow, containing_sheet_object);
                case CostItems.SACE:
                    return new ScheduleCostEstimate(xlRow, containing_sheet_object);
                case CostItems.I:
                    return new Input_Item(xlRow, containing_sheet_object);
                case CostItems.S:
                    return new Sum_Item(xlRow, containing_sheet_object);
                case CostItems.W:
                    return new WBS_Item(xlRow, containing_sheet_object);
                default:
                    throw new Exception("Unknown row type");
            }
            //I'm feeding this an xl row, picking up its type offset from its containing sheet specs

            throw new NotImplementedException();
        }

        public static void ExpandCorrelation()
        {
            Excel.Range selection = ThisAddIn.MyApp.Selection;
            SheetType sheetType = ExtensionMethods.GetSheetType(selection.Worksheet);
            if (sheetType != SheetType.Estimate && sheetType != SheetType.WBS) { ExtensionMethods.TurnOnUpdating(); return; }
            CostSheet sheetObj = CostSheet.ConstructFromXlCostSheet(selection.Worksheet);
            IEnumerable<Item> items = from Item item in sheetObj.Items where item.xlRow.Row == selection.Row select item;
            if (!items.Any()) { return; }
            Item selectedItem = items.First();
            CorrelationType correlType;

            //This needs done in the class so that I can know what type it is... Loading inputs with no subs doesn't work
            if (selection.Column == sheetObj.Specs.CostCorrel_Offset && selectedItem is IHasCostCorrelations)
            {
                correlType = CorrelationType.Cost;
            }
            else if (selection.Column == sheetObj.Specs.DurationCorrel_Offset && selectedItem is IHasDurationCorrelations)
            {
                correlType = CorrelationType.Duration;
            }
            else if (selection.Column == sheetObj.Specs.PhasingCorrel_Offset && selectedItem is IHasPhasingCorrelations)
            {
                correlType = CorrelationType.Phasing;
            }
            else
            {
                //Probably a misclick
                return;
                //correlType = CorrelationType.Null;
                //throw new Exception("Unknown Correlation Type");
            }

            ExtensionMethods.TurnOffUpdating();
            switch (correlType)
            {
                case CorrelationType.Cost:
                    if (selectedItem.CanExpand(correlType))
                        ((IHasSubs)selectedItem).Expand(correlType);
                    break;
                case CorrelationType.Duration:
                    if (selectedItem.CanExpand(correlType))
                        ((IHasSubs)selectedItem).Expand(correlType);
                    break;
                case CorrelationType.Phasing:
                    if (selectedItem.CanExpand(correlType))
                        ((IHasSubs)selectedItem).Expand(correlType);
                    break;
                case CorrelationType.Null:      //Not selecting a correlation column
                    return;
                default:
                    throw new Exception("Unknown correlation expand issue");
            }

            ExtensionMethods.TurnOnUpdating();
        }

    }
}
