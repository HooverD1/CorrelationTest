using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CorrelationTest
{
    public struct UniqueID
    {
        private const int UniqueID_Fields = 2;
        private const int SheetName_Placement = 0;
        public string SheetName { get; }
        private const int FieldName_Placement = 1;
        public string FieldName { get; }
        private string Postfix { get; }
        //private string WBSLevel { get; set; }
        public string Value { get; }

        public UniqueID(string FullID)
        {
            this.Postfix = null;
            this.Value = FullID;
            Dictionary<string, string> ID_Components = ParseID(this.Value);
            this.SheetName = ID_Components["SheetName"];
            this.FieldName = ID_Components["FieldName"];
        }

        public UniqueID(string SheetName, string FieldName, string Postfix = null)
        {
            StringBuilder sb = new StringBuilder();
            this.SheetName = SheetName;
            this.FieldName = FieldName;
            this.Postfix = Postfix;
            for (int i = 0; i < UniqueID_Fields; i++)
            {
                switch (i)
                {
                    case SheetName_Placement:
                        sb.Append(SheetName);
                        break;
                    case FieldName_Placement:
                        if (Postfix == null)
                            sb.Append($"{FieldName}");
                        else
                            sb.Append($"{FieldName} ({this.Postfix})");
                        break;
                    default:
                        break;
                }
                if (i < UniqueID_Fields - 1)
                    sb.Append("|");
            }
            this.Value = sb.ToString();
        }

        public bool Equals(UniqueID otherID)
        {
            if (this.Value == otherID.Value)
                return true;
            else
                return false;
        }

        private static Dictionary<string, string> ParseID(string Value)    //<property, value>
        {
            Dictionary<string, string> UniqueID_Properties = new Dictionary<string, string>();
            string[] valueSplit = Value.Split('|');
            UniqueID_Properties.Add("SheetName", valueSplit[0]);
            UniqueID_Properties.Add("FieldName", valueSplit[1]);
            return UniqueID_Properties;
        }

        public static void AutoFixUniqueIDs(List<IEstimate> Estimates)
        {
            for (int i = 0; i < Estimates.Count; i++)
            {
                //search the other estimates for the same name
                var duplicatedNames = (from IEstimate est in Estimates where Estimates[i].ID.Equals(est.ID) select est).ToArray();
                if (duplicatedNames.Count() > 1)
                {
                    for(int j = 0; j < duplicatedNames.Count(); j++)
                    {
                        duplicatedNames[j].Name = $"{duplicatedNames[j].Name} ({j+1})";
                        duplicatedNames[j].ID = new UniqueID(duplicatedNames[j].xlRow.Worksheet.Name, duplicatedNames[j].Name);
                    }
                }
                //Estimates[i].ID = new UniqueID(Estimates[i].xlRow.Worksheet.Name, Estimates[i].Name);
            }
            //Print out the new IDs
            foreach(IEstimate est in Estimates)
            {
                est.PrintName();
            }
        }

        public static UniqueID[] AutoFixUniqueIDs(UniqueID[] uniqueIDs)
        {
            if(uniqueIDs.Count() > 1)
            {
                for (int i = 0; i < uniqueIDs.Count(); i++)
                {
                    uniqueIDs[i] = new UniqueID(uniqueIDs[i].SheetName, uniqueIDs[i].FieldName, $"{i + 1}");
                }
            }                
            return uniqueIDs;
        }
    }
}
