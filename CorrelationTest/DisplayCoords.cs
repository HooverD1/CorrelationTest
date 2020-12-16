﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace CorrelationTest
{
    public class DisplayCoords
    {
        public int ID_Offset { get; set; }
        public int Level_Offset { get; set; }
        public int Type_Offset { get; set; }
        public int Name_Offset { get; set; }
        public int InputCorrel_Offset { get; set; }
        public int Distribution_Offset { get; set; }
        public int Dollar_Offset { get; set; }

        public static DisplayCoords ConstructDisplayCoords(SheetType displaySheet)
        {
            switch (displaySheet)
            {
                case SheetType.WBS:
                    return new DisplayCoords_WBS();
                case SheetType.Estimate:
                    return new DisplayCoords_Estimate();
                default:
                    throw new Exception("Invalid sheet type");
            }
        }
    }

    public class DisplayCoords_WBS : DisplayCoords
    {
        private const int WBS_ID_Offset = 1;
        private const int WBS_Level_Offset = 2;
        private const int WBS_Type_Offset = 3;
        private const int WBS_Name_Offset = 4;
        private const int WBS_InputCorrel_Offset = 5;
        private const int WBS_Distribution_Offset = 6;
        private const int WBS_Dollar_Offset = 12;

        public DisplayCoords_WBS()
        {
            this.ID_Offset = WBS_ID_Offset;
            this.Level_Offset = WBS_Level_Offset;
            this.Type_Offset = WBS_Type_Offset;
            this.Name_Offset = WBS_Name_Offset;
            this.InputCorrel_Offset = WBS_InputCorrel_Offset;
            this.Distribution_Offset = WBS_Distribution_Offset;
            this.Dollar_Offset = WBS_Dollar_Offset;
        }
    }

    public class DisplayCoords_Estimate : DisplayCoords
    {
        private const int Est_ID_Offset = 1;
        private const int Est_Level_Offset = 2;
        private const int Est_Type_Offset = 3;
        private const int Est_Name_Offset = 4;
        private const int Est_InputCorrel_Offset = 5;
        private const int Est_Distribution_Offset = 6;

        public DisplayCoords_Estimate()
        {
            this.ID_Offset = Est_ID_Offset;
            this.Level_Offset = Est_Level_Offset;
            this.Type_Offset = Est_Type_Offset;
            this.Name_Offset = Est_Name_Offset;
            this.InputCorrel_Offset = Est_InputCorrel_Offset;
            this.Distribution_Offset = Est_Distribution_Offset;
        }
    }
}