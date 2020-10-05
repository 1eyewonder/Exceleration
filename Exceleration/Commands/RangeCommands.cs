using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;
using Exceleration.Helpers;
using Exceleration.Helpers.Enums;
using Exceleration.Helpers.Extensions;

namespace Exceleration.Commands
{
    public class RangeCommands
    {
        public const string AddNamedRange = "ADD NAMED RANGE";
        public const string SetNamedRange = "SET NAMED RANGE";
        public const string RemoveNamedRange = "REMOVE NAMED RANGE";
        public const string RenameRange = "RENAME RANGE";
        public const string DeleteRangeContents = "DELETE RANGE CONTENTS";

        public void AddWorksheetNamedRange(Excel.Worksheet worksheet, string name, string range)
        {
            worksheet.CreateNamedRange(name, range);
        }

        public void AddWorkbookNamedRange(Excel.Workbook workbook, string name, string range)
        {
            workbook.CreateNamedRange(name, range);
        }

        public void SetNamedRangeCommand(Excel.Worksheet worksheet, string name, string range)
        {
            // If named range exists
            if (worksheet.NamedRangeExists(name))
            {
                // If valid range
                if (worksheet.IsRange(range))
                {
                    //worksheet.Range[name].Range[range]; 
                    worksheet.Names.Item(worksheet.Range[$"{name}"], Type.Missing, worksheet.Range[range]);
                    
                }
                else
                {
                    throw new ArgumentException($"Range entered {range} is not valid");
                }
            }
            else
            {
                throw new ArgumentException($"Named range {name} does not exist");
            }
        }

        public void RemoveNamedRangeCommand(Excel.Worksheet worksheet, string name)
        {
            worksheet.DeleteRangeContents(name);
        }

        public void RenameWorkbookRange(Excel.Workbook workbook, string oldName, string newName)
        {
            workbook.RenameRange(oldName, newName);
        }

        public void RenameWorksheetRange(Excel.Worksheet worksheet, string oldName, string newName)
        {
            worksheet.RenameRange(oldName, newName);
        }

        public void DeleteRangeContentsCommand(Excel.Worksheet worksheet, string name, ReferenceEnum referenceType = ReferenceEnum.ByName)
        {
            if (referenceType == ReferenceEnum.ByName)
            {
                worksheet.DeleteRangeContents(name);
            }
            else if (referenceType == ReferenceEnum.ByIndex)
            {
                worksheet.DeleteRangeContents(name, false);
            }         
        }
    }
}
