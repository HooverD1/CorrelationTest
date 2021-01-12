using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace CorrelationTest
{
    public class UniqueID       //UniqueID form: SheetType | CreatedDateTime
    {
        private const int SheetType_Placement = 1;
        public string SheetType { get; set; }
        private const int Created_Placement = 3; 
        public string Created { get; set; }
        private const int ProgramTag_Placement = 0;
        public string ProgramTag { get; set; }
        private const int User_Placement = 2;
        public string User { get; set; }
        protected const char Delimiter = '|';
        protected const char Delimiter2 = '.';
        public string ID { get; set; }

        public static UniqueID ConstructFromExisting(string existingID)
        {
            UniqueID returnID = new UniqueID();
            returnID.ID = existingID;
            Dictionary<string, string> ID_Components = ParseID(returnID.ID);
            returnID.SheetType = ID_Components["SheetType"];
            returnID.Created = ID_Components["Created"];
            returnID.User = ID_Components["User"];
            returnID.ProgramTag = ID_Components["ProgramTag"];
            return returnID;
        }

        public static UniqueID ConstructNew(string prefix, string created = null)
        {
            UniqueID returnID = new UniqueID();
            returnID.SheetType = prefix;
            if (created == null)
                returnID.Created = UniqueID.Timestamp();
            else
                returnID.Created = created;
            returnID.ID = returnID.CreateID(new Dictionary<string, string>() {  { "SheetType", returnID.SheetType },
                                                                                { "Created", returnID.Created },
                                                                                { "User", ThisAddIn.MyApp.UserName },
                                                                                { "ProgramTag", "DH"} });
            return returnID;
        }

        public void PrintToCell(Excel.Range xlUniqueID)
        {
            xlUniqueID.Value = this.ID;
        }

        public void RefreshID()
        {
            this.ID = CreateID(new Dictionary<string, string>() { { "SheetType", this.SheetType },
                                                                  { "Created", this.Created },
                                                                  { "User", ThisAddIn.MyApp.UserName },
                                                                  { "ProgramTag", this.ProgramTag} });
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
                    case Created_Placement:
                        sb.Append(ParamDict["Created"]);
                        break;
                    case User_Placement:
                        sb.Append(ParamDict["User"]);
                        break;
                    case ProgramTag_Placement:
                        sb.Append(ParamDict["ProgramTag"]);
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
            UniqueID_Properties.Add("Created", valueSplit[Created_Placement]);
            UniqueID_Properties.Add("User", valueSplit[User_Placement]);
            UniqueID_Properties.Add("ProgramTag", valueSplit[ProgramTag_Placement]);
            return UniqueID_Properties;
        }

        private static string Timestamp()
        {
            var returnval = $"{DateTime.Now.ToUniversalTime().ToString("ddMMyy")}{DateTime.Now.ToUniversalTime().ToString("HH:mm")}";
            return returnval;
        }

        public static bool Validate(string uidString)
        {
            try
            {
                string[] uidValues = uidString.Split('|');
                if (uidValues.Length != 3)
                    return false;
                return true;
            }
            catch(Exception)
            {
                return false;
            }
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
