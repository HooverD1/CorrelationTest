using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace CorrelationTest
{
    public class UniqueID
    {
        private const int SheetType_Placement = 0;
        public string SheetType { get; }
        private const int Name_Placement = 1;
        public string Name { get; }
        private const int Created_Placement = 2; 
        public string Created { get; }
        protected const char Delimiter = '|';
        protected const char Delimiter2 = '.';
        public string ID { get; set; }

        public UniqueID(string FullID)
        {
            this.ID = FullID;
            Dictionary<string, string> ID_Components = ParseID(this.ID);
            this.SheetType = ID_Components["SheetType"];
            this.Name = ID_Components["Name"];
            this.Created = ID_Components["Created"];
            
        }

        public UniqueID(string SheetType, string Name, string Created = null)
        {
            this.SheetType = SheetType;
            this.Name = Name;
            if (Created == null)
                this.Created = UniqueID.Timestamp();
            else
                this.Created = Created;
            this.ID = CreateID(new Dictionary<string, string>() { { "SheetType", this.SheetType },
                                                                            { "Name", this.Name },
                                                                            { "Created", this.Created } });
        }

        public UniqueID(Dictionary<string, string> ID_Components)
        {
            this.SheetType = string.Empty;
            this.Name = string.Empty;
            this.Created = string.Empty;
            if (ID_Components.ContainsKey("SheetType"))
                this.SheetType = ID_Components["SheetType"];
            if (ID_Components.ContainsKey("Name"))
                this.Name = ID_Components["Name"];
            if (ID_Components.ContainsKey("Created"))
                this.Created = ID_Components["Created"];
            else
                this.Created = UniqueID.Timestamp();
            if (string.IsNullOrEmpty(this.SheetType) || string.IsNullOrEmpty(this.Name))
                throw new Exception("Malformed dictionary parameter");
            else
                this.ID = CreateID(ID_Components);
        }

        public void PrintToCell(Excel.Range xlUniqueID)
        {
            xlUniqueID.Value = this.ID;
        }

        private string CreateID(Dictionary<string, string> ParamDict)
        {
            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < ParamDict.Count; i++)
            {
                switch (i)
                {
                    case SheetType_Placement:
                        sb.Append(ParamDict["SheetType"]);
                        break;
                    case Name_Placement:
                        sb.Append($"{ParamDict["Name"]}");
                        break;
                    case Created_Placement:
                        sb.Append(ParamDict["Created"]);
                        break;
                    default:
                        break;
                }
                if (i < ParamDict.Count - 1)
                    sb.Append(Delimiter);
            }
            return sb.ToString();
        }

        public bool Equals(UniqueID otherID)
        {
            if (this.ID == otherID.ID)
                return true;
            else
                return false;
        }

        private static Dictionary<string, string> ParseID(string Value)    //<property, value>
        {
            Dictionary<string, string> UniqueID_Properties = new Dictionary<string, string>();
            string[] valueSplit = Value.Split(Delimiter);
            UniqueID_Properties.Add("SheetType", valueSplit[SheetType_Placement]);
            UniqueID_Properties.Add("Name", valueSplit[Name_Placement]);
            UniqueID_Properties.Add("Created", valueSplit[Created_Placement]);
            return UniqueID_Properties;
        }

        private static string Timestamp()
        {
            var returnval = $"{DateTime.Now.ToUniversalTime().ToString("ddMMyy")}{DateTime.Now.ToUniversalTime().ToString("HH:mm")}";
            return returnval;
        }

        //public static void AutoFixUniqueIDs(List<IEstimate> Estimates)
        //{            
        //    for (int i = 0; i < Estimates.Count; i++)
        //    {
        //        //search the other estimates for the same name

        //        var duplicatedNames = (from IEstimate est in Estimates where Estimates[i].ID.Equals(est.ID) select est).ToArray();
        //        if (duplicatedNames.Count() > 1)
        //        {
        //            for(int j = 0; j < duplicatedNames.Count(); j++)
        //            {
        //                duplicatedNames[j].Name = $"{duplicatedNames[j].Name} ({j+1})";
        //                duplicatedNames[j].ID = new UniqueID(duplicatedNames[j].xlRow.Worksheet.Name, duplicatedNames[j].Name);
        //            }
        //        }
        //        //Estimates[i].ID = new UniqueID(Estimates[i].xlRow.Worksheet.Name, Estimates[i].Name);
        //    }
        //    //Print out the new IDs
        //    foreach(IEstimate est in Estimates)
        //    {
        //        est.PrintName();
        //    }
        //}

        //public static UniqueID[] AutoFixUniqueIDs(UniqueID[] uniqueIDs)
        //{
        //    if(uniqueIDs.Count() > 1)
        //    {
        //        for (int i = 0; i < uniqueIDs.Count(); i++)
        //        {
        //            uniqueIDs[i] = new UniqueID(uniqueIDs[i].SheetType, uniqueIDs[i].Name, $"{i + 1}");
        //        }
        //    }                
        //    return uniqueIDs;
        //}

        
    }
}
