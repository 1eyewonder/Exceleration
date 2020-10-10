using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using Exceleration.Helpers.Extensions;

namespace Exceleration.Commands
{
    public class WorksheetCommands
    {
        public const string AddColumn = "ADD COLUMN";
        public const string AddRow = "ADD ROW";

        public void AddColumnCommand(Excel.Worksheet worksheet, string range)
        {
            worksheet.AddColumn(range);
        }

        public void AddRowCommand(Excel.Worksheet worksheet, string range)
        {
            worksheet.AddRow(range);
        }
    }
}
