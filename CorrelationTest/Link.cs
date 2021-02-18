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
        public class Link
        {
            public string Address { get; }
            public Excel.Range LinkSource { get; }

            public Link(Excel.Range linkSource)
            {
                this.LinkSource = linkSource;
                this.Address = $"={linkSource.Address[External: true]}";
            }

            public Link(string linkSourceAddress)
            {
                //parse the linkSourceAddress into sheet and address
                string[] splitAddress = ParseAddress(linkSourceAddress);
                var xlSourceSheets = from Excel.Worksheet sheet in ThisAddIn.MyApp.Worksheets
                                     where sheet.Name == splitAddress[0]    //should probably come up with some kind of link ID in case someone deletes a sheet and renames another the same?
                                     select sheet;
                Excel.Worksheet xlSourceSheet;
                if (xlSourceSheets.Any())
                    xlSourceSheet = xlSourceSheets.First();
                else
                    throw new Exception();      //someone deleted the sheet!! 
                this.LinkSource = xlSourceSheet.Range[splitAddress[1]];
                this.Address = $"={this.LinkSource.Address[External: true]}";
            }

            public void PrintToSheet(Excel.Range linkTarget, string linkFormat = "\"Link\";;;\"LINK\"")
            {
                linkTarget.Value = this.LinkSource.Address[External: true];
                linkTarget.WrapText = false;
                linkTarget.NumberFormat = linkFormat;
                linkTarget.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            }
            private string[] ParseAddress(string linkSourceAddress)
            {
                string[] unbook = linkSourceAddress.Split(']');
                return unbook[1].Split('!');
            }


            
        }
    }
}
