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
        public class CorrelationString_IT : CorrelationString
        {
            public Triple InputTriple { get; set; }
            public CorrelationString_IT(string correlString)
            {
                this.Value = ExtensionMethods.CleanStringLinebreaks(correlString);
                string triple = this.Value.Split('&')[2];
                this.InputTriple = new Triple(this.GetParentID().ID, triple);
            }

            public CorrelationString_IT(Triple it, int subs, string parent_id)        //build a triple string out of a triple
            {
                this.InputTriple = it;
                StringBuilder sb = new StringBuilder();
                sb.Append($"{subs},IT");
                sb.AppendLine();
                sb.Append(parent_id);
                sb.AppendLine();
                sb.Append(it.ToString());
                this.Value = ExtensionMethods.CleanStringLinebreaks(sb.ToString());
            }

            public override void PrintToSheet(Excel.Range xlCell)
            {
                xlCell.Value = this.Value;
                xlCell.NumberFormat = "\"In Correl\";;;\"IN_CORREL\"";
                xlCell.EntireColumn.ColumnWidth = 10;
            }
        }
    }
}
