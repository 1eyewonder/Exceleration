using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace Exceleration.Helpers
{
    public static class WorkbookHelper
    {
        /// <summary>
        /// Automaps object properties to sheet specific named ranges within an Excel workbook
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="workbook">Excel workbook values are being written to</param>
        /// <param name="obj">Object values are being read from</param>
        /// <param name="skippedSheets">Sheets in the target workbook the user wants skipped</param>
        /// <param name="skippedProperties">Properties in the object the user wants skipped</param>
        public static void ExcelWorkbookMapper<T>(Excel.Workbook workbook, T obj, List<Excel.Worksheet> skippedSheets = null, List<string> skippedProperties = null)
        {
            //If optional parameters aren't set, set blank objects to prevent null object error
            if (skippedSheets == null) { skippedSheets = new List<Excel.Worksheet>(); }
            if (skippedProperties == null) { skippedProperties = new List<string>(); }

            //Loops through each worksheet in the target workbook
            foreach (Excel.Worksheet sheet in workbook.Worksheets)
            {
                //Skips any sheets included in the parameter list
                if (skippedSheets.Contains(sheet)) { continue; }

                WorksheetHelper.ExcelWorksheetMapper(sheet, obj, skippedProperties);
            }
        }
    }
}
