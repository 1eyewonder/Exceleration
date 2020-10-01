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
        /// Creates a new worksheet in the given workbook
        /// </summary>
        /// <param name="workbook">Target workbook</param>
        /// <param name="name">New sheet name</param>
        public static Excel.Worksheet CreateNewWorksheet(this Excel.Workbook workbook, string name = "NewSheet")
        {
            // If worksheet name already exists
            if (workbook.WorkSheetExists(name))
            {
                throw new ArgumentException("Worksheet name already exists");
            }
 
            // Adds worksheet
            Excel.Worksheet newWorksheet = (Excel.Worksheet)workbook.Worksheets.Add();
            newWorksheet.Name = $"{name}";

            return newWorksheet;
        }

        /// <summary>
        /// Moves worksheet to desired location within the current workbook
        /// </summary>
        /// <param name="worksheet">Worksheet being moved</param>
        /// <param name="workbook">Active workbook the worksheet exists in</param>
        /// <param name="positional">Desired position of worksheet. Default position is the end of the workbook.</param>
        /// <param name="referenceName">Name of worksheet position is relative to. Default worksheet is 'Commands"</param>
        /// <param name="referenceType">If relative worksheet is referenced through name or index</param>
        public static void MoveWorksheet(this Excel.Worksheet worksheet, 
            Excel.Workbook workbook,
            PositionalEnum positional = PositionalEnum.AtEnd,
            string referenceName = "Commands",
            ReferenceEnum referenceType = ReferenceEnum.ByName)
        {
            // Gets worksheet count to prep for positional placement, if needed
            var sheetCount = workbook.Worksheets.Count;
            int indexValue = 1;
            bool isAnInteger = false;
            
            if (referenceType == ReferenceEnum.ByIndex)
            {
                // Test if sheet name/index is an integer value
                isAnInteger = int.TryParse(referenceName, out indexValue);

                if (isAnInteger == false)
                {
                    throw new ArgumentException("Index value given is either not a string or is not an integer");
                }

                // Check index value exists within the current workbook
                if (indexValue <= 0 || indexValue > sheetCount)
                {
                    throw new ArgumentException("Index value is out of range");
                }
            }
            else
            {
                // If relative worksheet name is needed
                if (positional != PositionalEnum.AtBeginning && positional != PositionalEnum.AtEnd)
                {
                    // If relative worksheet does not exist
                    if (!workbook.WorkSheetExists(referenceName))
                    {
                        throw new ArgumentException("Relative worksheet name does not exist");
                    }
                }               
            }          

            // Moves worksheet to desired position
            switch (positional)
            {
                case PositionalEnum.AtBeginning:
                    worksheet.Move(workbook.Worksheets[1]);
                    break;

                case PositionalEnum.AtEnd:
                    worksheet.Move(After: workbook.Worksheets[sheetCount]);
                    break;

                case PositionalEnum.After:
                    if (referenceType == ReferenceEnum.ByIndex && isAnInteger == true)
                    {
                        worksheet.Move(After: workbook.Worksheets[indexValue]);
                    }
                    else
                    {
                        worksheet.Move(After: workbook.Worksheets[$"{referenceName}"]);
                    }
                    
                    break;

                case PositionalEnum.Before:
                    if (referenceType == ReferenceEnum.ByIndex && isAnInteger == true)
                    {
                        worksheet.Move(workbook.Worksheets[indexValue]);
                    }
                    else
                    {
                        worksheet.Move(workbook.Worksheets[$"{referenceName}"]);
                    }
                        
                    break;
            }
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
        /// <param name="name">Name of worksheet to be copied</param>
        /// <param name="newName">New worksheet name</param>
        public static void CopySheet(this Excel.Workbook workbook, string name, string newName = "NewSheet")
        {
            Excel.Worksheet worksheet = null;

            if (workbook.WorkSheetExists(name))
            {
                if (!workbook.WorkSheetExists(newName))
                {
                    worksheet = ((Excel.Worksheet)workbook.Worksheets[name]);
                    worksheet.Copy(Type.Missing, worksheet);
                    ((Excel.Worksheet)workbook.Worksheets[worksheet.Index + 1]).Name = newName;
                }
                else
                {
                    throw new ArgumentException("New name already has a worksheet with the same name");
                }

            }
            else
            {
                throw new ArgumentException("Worksheet does not exist");
            }
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
