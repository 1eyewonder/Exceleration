using Exceleration.Helpers.Enums;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace Exceleration.Helpers.Extensions
{
    public static class ExcelWorkbookExtensions
    {
        #region Worksheets
        /// <summary>
        /// Checks if worksheet with given name exists
        /// </summary>
        /// <param name="workbook">Target workbook</param>
        /// <param name="name">Worksheet name</param>
        /// <returns></returns>
        public static bool WorksheetExists(this Excel.Workbook workbook, string name)
        {
            var worksheets = workbook.GetWorksheets();
            var worksheet = worksheets.FirstOrDefault(x => x.Name == name);

            if (worksheet == null)
            {
                return false;
            }
            else
            {
                return true;
            }
        }

        /// <summary>
        /// Checks if worksheet with given index exists
        /// </summary>
        /// <param name="workbook">Target workbook</param>
        /// <param name="index">Worksheet index</param>
        /// <returns></returns>
        public static bool WorksheetExists(this Excel.Workbook workbook, int index)
        {
            int worksheetCount = workbook.Worksheets.Count;

            if (index < 0 || index > worksheetCount)
            {
                return false;
            }
            else
            {
                return true;
            }
        }

        /// <summary>
        /// Creates a new worksheet in the given workbook
        /// </summary>
        /// <param name="workbook">Target workbook</param>
        /// <param name="name">New sheet name</param>
        public static Excel.Worksheet CreateNewWorksheet(this Excel.Workbook workbook, string name = "NewSheet")
        {
            // If worksheet name already exists
            if (workbook.WorksheetExists(name))
            {
                throw new ArgumentException($"Worksheet name {name} already exists");
            }

            // Adds worksheet
            Excel.Worksheet newWorksheet = (Excel.Worksheet)workbook.Worksheets.Add();
            newWorksheet.Name = $"{name}";

            return newWorksheet;
        }

        /// <summary>
        /// Make the worksheet with the given name the active sheet in the workbook
        /// </summary>
        /// <param name="workbook">Target workbook</param>
        /// <param name="name">Worksheet name</param>
        /// <returns></returns>
        public static void ActivateSheet(this Excel.Workbook workbook, string name)
        {
            var worksheet = workbook.GetWorksheet(name);
            worksheet.Activate();
        }

        /// <summary>
        /// Make the worksheet with the given index the active sheet in the workbook
        /// </summary>
        /// <param name="workbook">Target workbook</param>
        /// <param name="index">Worksheet index</param>
        public static void ActivateSheet(this Excel.Workbook workbook, int index)
        {
            var worksheet = workbook.GetWorksheet(index);
            worksheet.Activate();
        }

        /// <summary>
        /// Copies Excel worksheet with given name and moves it to the end of the workbook
        /// </summary>
        /// <param name="workbook">Target workbook</param>
        /// <param name="name">Name of worksheet to be copied</param>
        /// <param name="newName">New worksheet name</param>
        public static void CopySheet(this Excel.Workbook workbook, string name, string newName = "NewSheet")
        {
            Excel.Worksheet worksheet = null;

            if (workbook.WorksheetExists(name))
            {
                if (!workbook.WorksheetExists(newName))
                {
                    worksheet = ((Excel.Worksheet)workbook.Worksheets[name]);
                    worksheet.Copy(Type.Missing, worksheet);
                    ((Excel.Worksheet)workbook.Worksheets[worksheet.Index + 1]).Name = newName;
                }
                else
                {
                    throw new ArgumentException($"New name {newName} already has a worksheet with the same name");
                }

            }
            else
            {
                throw new ArgumentException($"Worksheet {name} does not exist");
            }
        }

        /// <summary>
        /// Copies Excel worksheet with the given index and moves it to the end of the workbook
        /// </summary>
        /// <param name="workbook">Target workbook</param>
        /// <param name="index">Index of worksheet to be copied</param>
        /// <param name="newName">New worksheet name</param>
        public static void CopySheet(this Excel.Workbook workbook, int index, string newName = "NewSheet")
        {
            Excel.Worksheet worksheet = null;

            if (workbook.WorksheetExists(index))
            {
                if (!workbook.WorksheetExists(newName))
                {
                    worksheet = workbook.Worksheets[index];
                    worksheet.Copy(Type.Missing, worksheet);
                    ((Excel.Worksheet)workbook.Worksheets[worksheet.Index + 1]).Name = newName;
                }
                else
                {
                    throw new ArgumentException($"New name {newName} already has a worksheet with the same name");
                }
            }
            else
            {
                throw new ArgumentException($"Worksheet {index} does not exist");
            }
        }

        /// <summary>
        /// Renames worksheet
        /// </summary>
        /// <param name="workbook">Target workbook</param>
        /// <param name="oldName">Worksheet to be renamed</param>
        /// <param name="newName">New sheet name</param>
        public static void RenameSheet(this Excel.Workbook workbook, string oldName, string newName = "NewSheet")
        {
            Excel.Worksheet worksheet = null;

            if (workbook.WorksheetExists(oldName))
            {
                if (!workbook.WorksheetExists(newName))
                {
                    worksheet = workbook.GetWorksheet(oldName);
                    worksheet.Name = newName;
                }
                else
                {
                    throw new ArgumentException($"New name {newName} already has a worksheet with the same name");
                }
            }
            else
            {
                throw new ArgumentException($"Worksheet {oldName} does not exist");
            }
        }

        /// <summary>
        /// Renames worksheet
        /// </summary>
        /// <param name="workbook">Target workbook</param>
        /// <param name="index">Worksheet to be renamed</param>
        /// <param name="newName">New sheet name</param>
        public static void RenameSheet(this Excel.Workbook workbook, int index, string newName = "NewSheet")
        {
            Excel.Worksheet worksheet = null;

            if (workbook.WorksheetExists(index))
            {
                if (!workbook.WorksheetExists(newName))
                {
                    worksheet = workbook.GetWorksheet(index);
                    worksheet.Name = newName;
                }
                else
                {
                    throw new ArgumentException($"New name {newName} already has a worksheet with the same name");
                }
            }
            else
            {
                throw new ArgumentException($"Worksheet {index} does not exist");
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
        /// Returns worksheet in the workbook with the given name
        /// </summary>
        /// <param name="workbook">Target workbook</param>
        /// <param name="worksheetName">Worksheet name</param>
        /// <returns></returns>
        public static Excel.Worksheet GetWorksheet(this Excel.Workbook workbook, string worksheetName)
        {
            if (workbook.WorksheetExists(worksheetName))
            {
                return workbook.Worksheets[$"{worksheetName}"];
            }
            else
            {
                throw new ArgumentException($"Worksheet name [{worksheetName}] does not exist in this workbook.");
            }
        }

        /// <summary>
        /// Returns worksheet in the workbook with the given index
        /// </summary>
        /// <param name="workbook">Target workbook</param>
        /// <param name="worksheetIndex">Worksheet index</param>
        /// <returns></returns>
        public static Excel.Worksheet GetWorksheet(this Excel.Workbook workbook, int worksheetIndex)
        {
            if (workbook.WorksheetExists(worksheetIndex))
            {
                return workbook.Worksheets[worksheetIndex];
            }
            else
            {
                throw new ArgumentException($"Worksheet index {worksheetIndex} is out of range");
            }
        }

        /// <summary>
        /// Moves worksheet to desired location within the current workbook
        /// </summary>
        /// <param name="worksheet">Worksheet being moved</param>
        /// <param name="workbook">Active workbook the worksheet exists in</param>
        /// <param name="positional">Desired position of worksheet. Default position is the end of the workbook.</param>
        /// <param name="referenceName">Name of worksheet position is relative to</param>
        /// <param name="referenceType">If relative worksheet is referenced through name or index</param>
        public static void MoveWorksheet(this Excel.Workbook workbook, string worksheetName,
            PositionalEnum positional, string referenceName, ReferenceEnum referenceType = ReferenceEnum.ByName)
        {
            Excel.Worksheet worksheet = null;
            var sheetCount = workbook.Worksheets.Count;
            int indexValue = 1;
            bool isAnInteger = false;

            // Checks if target sheet exists
            if (!workbook.WorksheetExists(worksheetName))
            {
                throw new ArgumentException($"Worksheet [{worksheetName}] does not exist in the workbook");
            }

            worksheet = workbook.GetWorksheet(worksheetName);

            // If relative sheet for positioning is required
            if (positional != PositionalEnum.AtBeginning && positional != PositionalEnum.AtEnd)
            {
                if (string.IsNullOrEmpty(referenceName))
                {
                    throw new ArgumentException("A worksheet reference is required due to selected option");
                }
            }

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

                if (worksheet.Index == indexValue)
                {
                    throw new ArgumentException("Cannot move a worksheet when referencing its own index");
                }
            }
            else
            {
                // If relative worksheet name is needed
                if (positional != PositionalEnum.AtBeginning && positional != PositionalEnum.AtEnd)
                {
                    // If relative worksheet does not exist
                    if (!workbook.WorksheetExists(referenceName))
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
        /// Deletes worksheet with the given name
        /// </summary>
        /// <param name="workbook">Target workbook</param>
        /// <param name="name">Worksheet name</param>
        public static void DeleteWorksheet(this Excel.Workbook workbook, string name)
        {
            if (workbook.WorksheetExists(name))
            {
                ((Excel.Worksheet)workbook.Worksheets[name]).Delete();
            }
            else
            {
                throw new ArgumentException($"Worksheet name [{name}] does not exist in this workbook.");
            }
        }

        /// <summary>
        /// Deletes worksheet with the given index
        /// </summary>
        /// <param name="workbook">Target workbook</param>
        /// <param name="index">Worksheet index</param>
        public static void DeleteWorksheet(this Excel.Workbook workbook, int index)
        {
            if (workbook.WorksheetExists(index))
            {
                ((Excel.Worksheet)workbook.Worksheets[index]).Delete();
            }
            else
            {
                throw new ArgumentException($"Worksheet index [{index}] does not exist in this workbook.");
            }
        }
        #endregion

        #region Ranges
        /// <summary>
        /// Deletes any named ranges who have lost their cell references
        /// </summary>
        /// <param name="workbook">Target workbook</param>
        public static void ClearLostRanges(this Excel.Workbook workbook)
        {
            var ranges = workbook.Names;
            int i = 1;

            while (i <= ranges.Count)
            {
                var currentName = ranges.Item(i, Type.Missing, Type.Missing);
                var refersTo = currentName.RefersTo.ToString();
                if (refersTo.Contains("REF!"))
                {
                    ranges.Item(i, Type.Missing, Type.Missing).Delete();
                }
                else
                {
                    i++;
                }
            }
        }

        /// <summary>
        /// Checks if range exists in the workbook
        /// </summary>
        /// <param name="workbook">Target workbook</param>
        /// <param name="range">Range</param>
        /// <returns></returns>
        public static bool IsRange(this Excel.Workbook workbook, string range)
        {
            foreach (Excel.Worksheet worksheet in workbook.Worksheets)
            {
                try
                {
                    if (worksheet.IsRange(range))
                    {
                        return true;
                    }
                }
                catch
                {
                    continue;
                }
            }

            return false;
        }

        /// <summary>
        /// Checks if named range exists in the workbook
        /// </summary>
        /// <param name="workbook">Target workbook</param>
        /// <param name="name">Name of range</param>
        /// <param name="notOnWorksheet">Flags if user wants values unique to workbook scope</param>
        public static bool NamedRangeExists(this Excel.Workbook workbook, string name, bool notOnWorksheet = false)
        {
            // List of named ranges on a worksheet
            var sheetList = new List<string>();

            // List of named ranges in the entire workbook
            var workbookList = workbook.NamedRanges();

            // List of named ranges that exist only on the workbook scope
            var filteredList = new List<string>();

            if (notOnWorksheet == true)
            {
                // Adds all worksheet named ranges to sheet list (removes worksheet identifier i.e. Sheet1!, etc.)
                foreach (Excel.Worksheet sheet in workbook.Worksheets)
                {
                    sheetList.AddRange(sheet.NamedRanges());
                }

                // Only keeps unique workbook named ranges on filtered list
                filteredList = workbookList.Except(sheetList).ToList();
                return filteredList.Exists(x => x == name);
            }
            else
            {
                return workbookList.Exists(x => x == name);
            }
        }

        /// <summary>
        /// Returns a list of all named ranges that exist in the workbook
        /// </summary>
        /// <param name="workbook">Target workbook</param>
        /// <returns></returns>
        public static List<string> NamedRanges(this Excel.Workbook workbook)
        {
            var value = new List<string>();

            foreach (Excel.Name n in workbook.Names)
            {
                // Gets everything after the ! since Excel returns names such as 'Sheet1!Testing'
                value.Add(n.Name.Split('!').Last());
            }

            return value;
        }

        /// <summary>
        /// Creates a named range on the workbook scope
        /// </summary>
        /// <param name="workbook">Target workbook</param>
        /// <param name="name">Name of range being created</param>
        /// <param name="range">Target range</param>
        /// <param name="scope">Scope named range is to be created on</param>
        public static void CreateNamedRange(this Excel.Workbook workbook, string name, string range)
        {
            Excel.Worksheet worksheet = workbook.ActiveSheet;
            string worksheetName;

            // Checks named range exists
            if (workbook.NamedRangeExists(name))
            {
                throw new ArgumentException($"Name {name} already exists");
            }

            if (!workbook.IsRange(range))
            {
                throw new ArgumentException($"Range entered {range} is not a valid range for the workbook");
            }

            // If range exists on another worksheet
            if (range.Contains("!"))
            {
                worksheetName = range.Split('!').First();
                worksheet = workbook.GetWorksheet(worksheetName);
            }

            workbook.Names.Add(name, worksheet.Range[range]);
        }

        /// <summary>
        /// Gets name object from workbook
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="name"></param>
        /// <returns></returns>
        public static Excel.Name GetNamedRange(this Excel.Workbook workbook, string name)
        {
            foreach (Excel.Name item in workbook.Names)
            {
                if (item.ShortName() == name)
                {
                    return item;
                }
            }

            throw new ArgumentException($"Name {name} does not exist");
        }

        /// <summary>
        /// Renames a range in the workbook
        /// </summary>
        /// <param name="workbook">Target workbook</param>
        /// <param name="oldName">Old range name</param>
        /// <param name="newName">New range name</param>
        public static void RenameRange(this Excel.Workbook workbook, string oldName, string newName)
        {
            if (workbook.NamedRangeExists(oldName))
            {
                if (!workbook.NamedRangeExists(newName))
                {
                    var item = workbook.GetNamedRange(oldName);
                    item.Name = newName;
                }
                else
                {
                    throw new ArgumentException($"Name {newName} already exists");
                }
            }
            else
            {
                throw new ArgumentException($"Range {oldName} does not exist");
            }
        }

        /// <summary>
        /// Removes named range from workbook scope
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="name"></param>
        public static void RemoveNamedRange(this Excel.Workbook workbook, string name)
        {
            workbook.GetNamedRange(name).Delete();
        }

        /// <summary>
        /// Deletes the workbook's range contents
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="range"></param>
        /// <param name="isNamedRange"></param>
        public static void DeleteRangeContents(this Excel.Workbook workbook, string range, bool isNamedRange = true)
        {
            string newRange;
            string worksheetName;
            Excel.Worksheet worksheet;

            if (isNamedRange)
            {
                if (workbook.NamedRangeExists(range))
                {
                    workbook.GetNamedRange(range).RefersToRange.Cells.Delete();
                }
            }
            else
            {
                if (workbook.IsRange(range))
                {
                    if (range.Contains('!'))
                    {
                        worksheetName = range.Split('!').First();
                        newRange = range.Split('!').Last();
                        worksheet = workbook.GetWorksheet(worksheetName);
                        worksheet.DeleteRangeContents(newRange, false);
                    }
                }
            }
        }
        #endregion
    }
}
