using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.IO;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Excel;

namespace Exceleration.Helpers
{
    public static class WorksheetHelper
    {
        /// <summary>
        /// Checks if worksheet with given name exists
        /// </summary>
        /// <param name="workbook">Target workbook</param>
        /// <param name="name">Worksheet name</param>
        /// <returns></returns>
        public static bool WorkSheetExists(this Excel.Workbook workbook, string name)
        {
            var value = false;

            foreach (Excel.Worksheet s in workbook.Worksheets)
            {
                if (s.Name == name)
                {
                    value = true;

                    break;
                }
            }

            return value;
        }

        /// <summary>
        /// Returns a list of worksheets in the target workbook
        /// </summary>
        /// <param name="workbook"></param>
        /// <returns></returns>
        public static IList<Excel.Worksheet> GetWorksheets(this Excel.Workbook workbook)
        {
            var tempList = new List<Excel.Worksheet>();

            foreach (Excel.Worksheet sheet in workbook.Worksheets)
            {
                tempList.Add(sheet);
            }

            return tempList;
        }

        /// <summary>
        /// Make the worksheet with the given name the active sheet in the workbook
        /// </summary>
        /// <param name="workbook">Target workbook</param>
        /// <param name="name">Worksheet name</param>
        /// <returns></returns>
        public static bool ActivateSheet(this Excel.Workbook workbook, string name)
        {
            var workSheets = workbook.GetWorksheets();

            var workSheet = workSheets.FirstOrDefault(x => x.Name == name);

            if (workSheet == null)
            {
                return false;
            }

            workSheet.Activate();

            return true;
        }

        /// <summary>
        /// Checks if worksheet is empty
        /// </summary>
        /// <param name="worksheet">Target worksheet</param>
        /// <returns></returns>
        public static bool WorkSheetEmpty(this Excel.Worksheet worksheet)
        {
            return worksheet.Application.WorksheetFunction.CountA(worksheet.UsedRange) == 0 && worksheet.Shapes.Count == 0;
        }

        /// <summary>
        /// Copies Excel workseheet and moves it to the end of the workbook
        /// </summary>
        /// <param name="workbook">Target workbook</param>
        /// <param name="sheet">Worksheet to be copied</param>
        /// <param name="name">New worksheet name</param>
        public static void CopySheet(this Excel.Workbook workbook, Excel.Worksheet sheet, string name)
        {
            sheet.Copy(Type.Missing, sheet);

            Excel.Worksheet newsheet = workbook.Sheets[sheet.Index + 1];

            newsheet.Name = $"{name}";

            Excel.Worksheet newlocation = workbook.Sheets[workbook.Worksheets.Count];

            newsheet.Move(Type.Missing, newlocation);
        }

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
