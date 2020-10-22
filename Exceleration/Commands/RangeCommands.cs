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
using Exceleration.CoreHelpers;

namespace Exceleration.Commands
{
    public class RangeCommands
    {
        public const string AddNamedRange = "ADD NAMED RANGE";
        public const string SetNamedRange = "SET NAMED RANGE";
        public const string RemoveNamedRange = "REMOVE NAMED RANGE";
        public const string RenameRange = "RENAME RANGE";
        public const string DeleteRangeContents = "DELETE RANGE CONTENTS";
        public const string GetColumnRange = "GET COLUMN RANGE";
        public const string GetRowRange = "GET ROW RANGE";
        public const string GetDataSetRange = "GET DATASET RANGE";
        public const string GetHeaderCell = "GET HEADER CELL";

        public void AddWorkbookNamedRange(Excel.Workbook workbook, string name, string range)
        {
            workbook.CreateNamedRange(name, range);
        }

        public void AddWorksheetNamedRange(Excel.Worksheet worksheet, string name, string range)
        {
            worksheet.CreateNamedRange(name, range);
        }

        public void SetNamedRangeCommand(Excel.Workbook workbook, string name, string range)
        {
            if (workbook.NamedRangeExists(name))
            {
                if (workbook.IsRange(range))
                {
                    workbook.GetNamedRange(name).RefersToLocal = "=" + range;
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

        public void RemoveWorksheetNamedRange(Excel.Worksheet worksheet, string name)
        {
            worksheet.RemoveNamedRange(name);
        }

        public void RemoveWorkbookNamedRange(Excel.Workbook workbook, string name)
        {
            workbook.RemoveNamedRange(name);
        }

        public void RenameWorkbookRange(Excel.Workbook workbook, string oldName, string newName)
        {
            workbook.RenameRange(oldName, newName);
        }

        public void RenameWorksheetRange(Excel.Worksheet worksheet, string oldName, string newName)
        {
            worksheet.RenameRange(oldName, newName);
        }

        public void DeleteWorksheetRangeContents(Excel.Worksheet worksheet, string range, ReferenceEnum referenceType = ReferenceEnum.ByName)
        {
            if (referenceType == ReferenceEnum.ByName)
            {
                worksheet.DeleteRangeContents(range);
            }
            else if (referenceType == ReferenceEnum.ByIndex)
            {
                worksheet.DeleteRangeContents(range, false);
            }         
        }

        public void DeleteWorkbookRangeContents(Excel.Workbook workbook, string range, ReferenceEnum referenceType = ReferenceEnum.ByName)
        {
            if (referenceType == ReferenceEnum.ByName)
            {
                workbook.DeleteRangeContents(range);
            }
            else if (referenceType == ReferenceEnum.ByIndex)
            {
                workbook.DeleteRangeContents(range, false);
            }
        }
     
        public string GetColumnRangeCommand(Excel.Worksheet worksheet, string range)
        {
            Excel.Range theRange = worksheet.GetColumnRange(range);
            return theRange.Address;
        }
       
        public string GetRowRangeCommand(Excel.Worksheet worksheet, string range)
        {
            Excel.Range theRange = worksheet.GetRowRange(range);
            return theRange.Address;
        }

        public string GetDataSetRangeCommand(Excel.Worksheet worksheet, string range)
        {
            Excel.Range theRange = worksheet.GetDataSetRange(range);
            return theRange.Address;
        }

        public string GetHeaderCellCommand(Excel.Worksheet worksheet, string range, string headerText)
        {
            Excel.Range theRange = worksheet.GetHeaderCell(range, headerText);

            if (theRange != null)
            {
                return theRange.Address;
            }
            else
            {
                throw new Exception($"The header, {headerText}, could not be found in the desired row.");
            }
        }

        public void Test(Excel.Workbook workbook, Excel.Worksheet activeSheet, string objectMap, string dataRange)
        {
            Excel.Worksheet worksheet = workbook.GetWorksheet("Object Maps");
            var objectMapHelper = new ObjectMapHelper(worksheet);

            if (worksheet.NamedRangeExists(objectMap))
            {
                var dictionary = objectMapHelper.GetObjectMap(objectMap);
                //var item = worksheet.Range[range].ConvertToDataTable();
                

                if (activeSheet.IsRange(dataRange))
                {
                    var someTable = activeSheet.Range[dataRange].ConvertToDataTable(dictionary);
                    Console.WriteLine("test");
                }
            }
        }
    }    
}
