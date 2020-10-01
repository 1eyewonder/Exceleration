using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using Exceleration.Helpers;
using Exceleration.Helpers.Enums;

namespace Exceleration.Options
{
    public class SheetCommands
    {
        public const string AddSheet = "ADD SHEET";
        public const string DeleteSheet = "DELETE SHEET";
        public const string CopySheet = "COPY SHEET";
        public const string MoveSheet = "MOVE SHEET";

        public void AddSheetCommand(Excel.Workbook workbook, string name = "NewSheet")
        {
            var newWorksheet = workbook.CreateNewWorksheet(name);
        }

        public void DeleteSheetCommand(Excel.Workbook workbook, string name)
        {
            ((Excel.Worksheet)workbook.Worksheets[name]).Delete();
        }

        public void CopySheetCommand(Excel.Workbook workbook, string name, string newName = "NewSheet")
        {
            workbook.CopySheet(name, newName);
        }

        public void MoveWorksheetCommand(Excel.Workbook workbook, 
            string name,
            PositionalEnum positional = PositionalEnum.AtEnd,
            string referenceName = null,
            ReferenceEnum referenceType = ReferenceEnum.ByName)
        {
            if (workbook.WorkSheetExists(name))
            {
                Excel.Worksheet targetSheet = workbook.Worksheets[$"{name}"];
                targetSheet.MoveWorksheet(workbook, positional, referenceName, referenceType);
            }
            else
            {
                throw new ArgumentException("Worksheet does not exist");
            }
            
        }
    }
}
