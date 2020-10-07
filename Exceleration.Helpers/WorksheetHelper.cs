using Exceleration.Helpers.Enums;
using System;
using System.Collections.Generic;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace Exceleration.Helpers
{
    public static class WorksheetHelper
    {                   
        /// <summary>
        /// Automaps object properties to named ranges of an Excel worksheet
        /// </summary>
        /// <typeparam name="T">Object type to have properties mapped to Excel</typeparam>
        /// <param name="sheet">Worksheet being filled out</param>
        /// <param name="obj">Object property values are being read from</param>
        /// <param name="skippedProperties">Properties in the object the user wants skipped</param>
        /// Properties to be mapped are case sensitive. Property names that match named ranges on the sheet
        /// will be placed into their respective named ranges. I created this method with the intent to loosely couple
        /// worksheet cells so developers were less concerned with the evolution of a workbook. This of course will
        /// only work if people making the Excel sheets remember to rename ranges accordingly
        public static void ExcelWorksheetMapper<T>(Excel.Worksheet sheet, T obj, List<string> skippedProperties = null)
        {
            if (skippedProperties == null) { skippedProperties = new List<string>(); }

            //Iterates through each named range in the worksheet
            foreach (Excel.Name name in sheet.Names)
            {
                //Iterates through each properties in the object
                foreach (var prop in obj.GetType().GetProperties())
                {
                    //Skips any object properties in the parameter list
                    if (skippedProperties.Contains(prop.Name)) { continue; }

                    //If named range contains property name
                    var sheetName = sheet.Name + "!";
                    string rangeName = name.Name.Replace("'", string.Empty);
                    rangeName = rangeName.Replace(sheetName, string.Empty);
                    if (rangeName.Equals(prop.Name))
                    {
                        //Sets the range value to the property value
                        sheet.Range[prop.Name].Value = prop.GetValue(obj, null);
                    }
                }
            }
        }

        /// <summary>
        /// Gets the name of a column
        /// </summary>
        /// <param name="columnNumber">Column number</param>
        /// <returns></returns>
        public static string GetColumnName(int columnNumber)
        {
            const string letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
            string columnName = "";

            while (columnNumber > 0)
            {
                columnName = letters[(columnNumber - 1) % 26] + columnName;
                columnNumber = (columnNumber - 1) / 26;
            }

            return columnName;
        }
    }
}
