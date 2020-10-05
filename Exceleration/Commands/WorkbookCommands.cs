using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using Exceleration.Helpers;
using Exceleration.Helpers.Enums;
using Exceleration.Helpers.Extensions;

namespace Exceleration.Options
{
    public class WorkbookCommands
    {
        public const string AddSheet = "ADD SHEET";
        public const string DeleteSheet = "DELETE SHEET";
        public const string CopySheet = "COPY SHEET";
        public const string MoveSheet = "MOVE SHEET";
        public const string TargetSheet = "TARGET SHEET";
        public const string RenameSheet = "RENAME SHEET";

        public void AddSheetCommand(Excel.Workbook workbook, string name = "NewSheet")
        {
            var newWorksheet = workbook.CreateNewWorksheet(name);
        }

        public void DeleteSheetCommand(Excel.Workbook workbook, string name, 
            ReferenceEnum referenceType = ReferenceEnum.ByName)
        {
            bool isAnInteger = int.TryParse(name, out int indexValue);

            if (referenceType == ReferenceEnum.ByIndex && isAnInteger == true)
            {
                workbook.DeleteWorksheet(indexValue);
            }
            else
            {
                workbook.DeleteWorksheet(name);
            };
        }

        public void CopySheetCommand(Excel.Workbook workbook, string name, string newName, 
            ReferenceEnum referenceType = ReferenceEnum.ByName)
        {
            bool isAnInteger = int.TryParse(name, out int indexValue);

            if (referenceType == ReferenceEnum.ByIndex && isAnInteger == true)
            {
                workbook.CopySheet(indexValue, newName);
            }
            else
            {
                workbook.CopySheet(name, newName);
            }            
        }

        public void MoveWorksheetCommand(Excel.Workbook workbook, string worksheetName,
            PositionalEnum positional = PositionalEnum.AtEnd, string referenceName = null,
            ReferenceEnum referenceType = ReferenceEnum.ByName)
        {
            workbook.MoveWorksheet(worksheetName, positional, referenceName, referenceType);
        }

        public void TargetSheetCommand(Excel.Workbook workbook, string name, 
            ReferenceEnum referenceType = ReferenceEnum.ByName)
        {
            bool isAnInteger = int.TryParse(name, out int indexValue);

            if (referenceType == ReferenceEnum.ByIndex && isAnInteger == true)
            {
                workbook.ActivateSheet(indexValue);
            }
            else
            {
                workbook.ActivateSheet(name);
            }          
        }

        public void RenameSheetCommand(Excel.Workbook workbook, string oldName,
            string newName, ReferenceEnum referenceType = ReferenceEnum.ByName)
        {
            bool isAnInteger = int.TryParse(oldName, out int indexValue);

            if (referenceType == ReferenceEnum.ByIndex && isAnInteger == true)
            {
                workbook.RenameSheet(indexValue, newName);
            }
            else
            {
                workbook.RenameSheet(oldName, newName);
            }
        }
    }
}
