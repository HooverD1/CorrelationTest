﻿using System;
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
        public Excel.Range xlCorrelCell_Inputs { get; set; }
        public Excel.Range xlCorrelCell_Periods { get; set; }
        public CostSheet ContainingSheetObject { get; set; }

        protected Item(Excel.Range xlRow, CostSheet ContainingSheetObject)
        {
            this.xlRow = xlRow;
            this.ContainingSheetObject = ContainingSheetObject;
            this.xlTypeCell = xlRow.Cells[1, ContainingSheetObject.Specs.Type_Offset];
            this.xlNameCell = xlRow.Cells[1, ContainingSheetObject.Specs.Name_Offset];
            this.xlCorrelCell_Inputs = xlRow.Cells[1, ContainingSheetObject.Specs.InputCorrel_Offset];
            this.xlCorrelCell_Periods = xlRow.Cells[1, ContainingSheetObject.Specs.PhasingCorrel_Offset];
        }

        public static Item Construct(Excel.Range xlRow, CostSheet containing_sheet_object)
        {
            //What kind of sheet object is constructing it?
            //What kind of CostRow is this?
            int type_offset = containing_sheet_object.Specs.Type_Offset;
            string type_value = Convert.ToString(xlRow.Cells[1, type_offset].value);
            CostItems costRow_type = CostItems.Null;
            foreach(CostItems ci in Enum.GetValues(typeof(CostItems)))      //Get the cell value into an enum
            {
                if (type_value == ci.ToString())
                    costRow_type = ci;
            }
            switch (costRow_type)       //Construct subtypes based on the enum
            {
                case CostItems.CE:
                    return new Estimate_Item(xlRow, containing_sheet_object);
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
    }
}