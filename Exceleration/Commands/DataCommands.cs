using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using Exceleration.Helpers.Extensions;
using Exceleration.Options;

namespace Exceleration.Commands
{
    public class DataCommands
    {
        public const string SetValue = "SET VALUE";
        public const string FindAndReplace = "FIND AND REPLACE";

        public void FindAndReplaceCommand(Excel.Worksheet worksheet, string range, string oldText, string newText, string matchValue, bool matchCase)
        {
            if (worksheet.IsRange(range))
            {
                if (matchValue == MatchValueOptions.MatchAll)
                {
                    worksheet.Range[range].FindAndReplace(oldText, newText, true, matchCase);
                }
                else
                {
                    worksheet.Range[range].FindAndReplace(oldText, newText, false, matchCase);
                }              
            }
            else
            {
                throw new Exception($"The range, {range}, does not exist on the current worksheet, {worksheet.Name}");
            }
        }
    }
}
