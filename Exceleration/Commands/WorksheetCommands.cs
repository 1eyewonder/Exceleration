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
        public const string MoveColumn = "MOVE COLUMN";
        public const string MoveRow = "MOVE ROW";
        public const string DeleteColumn = "DELETE COLUMN";
        public const string DeleteRow = "DELETE ROW";

        public void AddColumnCommand(Excel.Worksheet worksheet, string range)
        {
            worksheet.AddColumn(range);
        }

        public void AddRowCommand(Excel.Worksheet worksheet, string range)
        {
            worksheet.AddRow(range);
        }

        public void MoveColumnCommand(Excel.Worksheet worksheet, string oldRange, string newRange)
        {
            worksheet.MoveColumn(oldRange, newRange);
        }

        public void MoveRowCommand(Excel.Worksheet worksheet, string oldRange, string newRange)
        {
            worksheet.MoveRow(oldRange, newRange);
        }      

        public void DeleteColumnCommand(Excel.Worksheet worksheet, string range)
        {
            worksheet.DeleteColumn(range);
        }

        public void DeleteRowCommand(Excel.Worksheet worksheet, string range)
        {
            worksheet.DeleteRow(range);
        }
    }
}
