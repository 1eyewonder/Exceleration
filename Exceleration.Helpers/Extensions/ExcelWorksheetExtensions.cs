using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace Exceleration.Helpers.Extensions
{
    public static class ExcelWorksheetExtensions
    {
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
        /// Adds column to the worksheet
        /// </summary>
        /// <param name="worksheet">Target worksheet</param>
        /// <param name="range">Range of a single cell on the sheet</param>
        public static void AddColumn(this Excel.Worksheet worksheet, string range)
        {
            if (string.IsNullOrEmpty(range))
            {
                throw new Exception($"Please enter a value for the range that is not null or empty");
            }

            if (worksheet.IsRange(range))
            {
                if (worksheet.Range[range].IsSingularCell())
                {
                    worksheet.Range[range].EntireColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight, Excel.XlInsertFormatOrigin.xlFormatFromLeftOrAbove);
                }
                else
                {
                    throw new Exception("Please ensure only one cell is reference to indicate the column");
                }
            }
            else
            {
                throw new Exception($"The range {range} is not a valid range on worksheet {worksheet.Name}");
            }
        }

        /// <summary>
        /// Deletes a column from the worksheet
        /// </summary>
        /// <param name="worksheet">Target worksheet</param>
        /// <param name="range">Singular cell range within the target column</param>
        public static void DeleteColumn(this Excel.Worksheet worksheet, string range)
        {
            if (string.IsNullOrEmpty(range))
            {
                throw new Exception($"Please enter a value for the range that is not null or empty");
            }

            if (worksheet.IsRange(range))
            {
                if (worksheet.Range[range].IsSingularCell())
                {
                    worksheet.Range[range].EntireColumn.Delete(Excel.XlDeleteShiftDirection.xlShiftToLeft);
                }
                else
                {
                    throw new Exception("Please ensure only one cell is reference to indicate the column");
                }
            }
            else
            {
                throw new Exception($"The range {range} is not a valid range on worksheet {worksheet.Name}");
            }
        }

        /// <summary>
        /// Moves a column within the current worksheet
        /// </summary>
        /// <param name="worksheet">Target worksheet</param>
        /// <param name="oldRange">Singular cell range within the target column</param>
        /// <param name="newRange">Singular cell range within the column the target is to be moved left of</param>
        public static void MoveColumn(this Excel.Worksheet worksheet, string oldRange, string newRange)
        {
            if (string.IsNullOrEmpty(oldRange) || string.IsNullOrEmpty(newRange))
            {
                throw new Exception($"The old or newly specified range is empty. Please specify which column is moving and where you would like to move it to.");
            }

            Excel.Range copyRange;
            Excel.Range insertRange;

            if (worksheet.IsRange(oldRange))
            {
                if (worksheet.Range[oldRange].IsSingularCell())
                {
                    if (worksheet.IsRange(newRange))
                    {
                        if (worksheet.Range[newRange].IsSingularCell())
                        {
                            copyRange = worksheet.Range[oldRange].EntireColumn;
                            insertRange = worksheet.Range[newRange].EntireColumn;
                            insertRange.Insert(Excel.XlInsertShiftDirection.xlShiftToRight, copyRange.Cut());
                        }
                        else
                        {
                            throw new Exception($"The column target location, {newRange}, does not specify one cell. Please verify one cell is selected in the desired column location.");
                        }
                    }
                    else
                    {
                        throw new Exception($"The column target location, {newRange}, is not a valid range. Please specify the range exists on the current worksheet and uses the correct formatting.");
                    }
                }
                else
                {
                    throw new Exception($"The targeted column range, {oldRange}, does not specify one cell. Please verify one cell is selected in the targeted column location.");
                }
            }
            else
            {
                throw new Exception($"The targeted column range, {oldRange}, is not a valid range. Please specify the range exists on the current worksheet and uses the correct formatting.");
            }
        }

        /// <summary>
        /// Moves a row within the current worksheet
        /// </summary>
        /// <param name="worksheet">Target worksheet</param>
        /// <param name="oldRange">Singular cell range within the target row</param>
        /// <param name="newRange">Singular cell range within the row the target is to be moved above</param>
        public static void MoveRow(this Excel.Worksheet worksheet, string oldRange, string newRange)
        {
            if (string.IsNullOrEmpty(oldRange) || string.IsNullOrEmpty(newRange))
            {
                throw new Exception($"The old or newly specified range is empty. Please specify which row is moving and where you would like to move it to.");
            }

            Excel.Range copyRange;
            Excel.Range insertRange;

            if (worksheet.IsRange(oldRange))
            {
                if (worksheet.Range[oldRange].IsSingularCell())
                {
                    if (worksheet.IsRange(newRange))
                    {
                        if (worksheet.Range[newRange].IsSingularCell())
                        {
                            copyRange = worksheet.Range[oldRange].EntireRow;
                            insertRange = worksheet.Range[newRange].EntireRow;
                            insertRange.Insert(Excel.XlInsertShiftDirection.xlShiftDown, copyRange.Cut());
                        }
                        else
                        {
                            throw new Exception($"The row target location, {newRange}, does not specify one cell. Please verify one cell is selected in the desired row location.");
                        }
                    }
                    else
                    {
                        throw new Exception($"The row target location, {newRange}, is not a valid range. Please specify the range exists on the current worksheet and uses the correct formatting.");
                    }
                }
                else
                {
                    throw new Exception($"The targeted row range, {oldRange}, does not specify one cell. Please verify one cell is selected in the targeted row location.");
                }
            }
            else
            {
                throw new Exception($"The targeted row range, {oldRange}, is not a valid range. Please specify the range exists on the current worksheet and uses the correct formatting.");
            }
        }

        /// <summary>
        /// Adds a row to the worksheet
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="range">Range of a single cell on the sheet</param>
        public static void AddRow(this Excel.Worksheet worksheet, string range)
        {
            if (string.IsNullOrEmpty(range))
            {
                throw new Exception($"Please enter a value for the range that is not null or empty");
            }
            if (worksheet.IsRange(range))
            {
                if (worksheet.Range[range].IsSingularCell())
                {
                    worksheet.Range[range].EntireRow.Insert(Excel.XlInsertShiftDirection.xlShiftDown, Excel.XlInsertFormatOrigin.xlFormatFromLeftOrAbove);
                }
                else
                {
                    throw new Exception("Please ensure only one cell is reference to indicate the row");
                }
            }
            else
            {
                throw new Exception($"The range {range} is not a valid range on worksheet {worksheet.Name}");
            }
        }

        /// <summary>
        /// Deletes a row from the worksheet
        /// </summary>
        /// <param name="worksheet">Target worksheet</param>
        /// <param name="range">Singular cell range within the target row</param>
        public static void DeleteRow(this Excel.Worksheet worksheet, string range)
        {
            if (string.IsNullOrEmpty(range))
            {
                throw new Exception($"Please enter a value for the range that is not null or empty");
            }

            if (worksheet.IsRange(range))
            {
                if (worksheet.Range[range].IsSingularCell())
                {
                    worksheet.Range[range].EntireRow.Delete(Excel.XlDeleteShiftDirection.xlShiftUp);
                }
                else
                {
                    throw new Exception("Please ensure only one cell is reference to indicate the row");
                }
            }
            else
            {
                throw new Exception($"The range {range} is not a valid range on worksheet {worksheet.Name}");
            }
        }

        #region Ranges
        /// <summary>
        /// Checks if range exists in the given worksheet scope
        /// </summary>
        /// <param name="workSheet">Target worksheet</param>
        /// <param name="name">Name of range user is checking</param>
        /// <returns></returns>
        public static bool NamedRangeExists(this Excel.Worksheet workSheet, string name)
        {
            return workSheet.NamedRanges().Exists(x => x == name);
        }

        /// <summary>
        /// Checks if cell reference represents a valid worksheet range
        /// </summary>
        /// <param name="worksheet">Target worksheet</param>
        /// <param name="range">Range</param>
        /// <returns></returns>
        public static bool IsRange(this Excel.Worksheet worksheet, string range)
        {
            try
            {
                var validRange = worksheet.Range[$"{range}", Type.Missing];
                return true;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Returns a list of all named ranges that exist on the target worksheet
        /// </summary>
        /// <param name="workSheet">Target worksheet</param>
        /// <returns></returns>
        public static List<string> NamedRanges(this Excel.Worksheet workSheet)
        {
            var value = new List<string>();

            foreach (Excel.Name n in workSheet.Names)
            {
                // Gets everything after the ! since Excel returns names such as 'Sheet1!Testing'
                value.Add(n.Name.Split('!').Last());
            }

            return value;
        }

        /// <summary>
        /// Creates a named range on the worksheet scope
        /// </summary>
        /// <param name="worksheet">Target worksheet</param>
        /// <param name="name">Name of range being created</param>
        /// <param name="range">Target range. Must be within the current worksheet</param>
        public static void CreateNamedRange(this Excel.Worksheet worksheet, string name, string range)
        {
            // If cell range isn't valid
            if (!worksheet.IsRange(range))
            {
                throw new ArgumentException($"Range entered {range} is not a valid range or does not exist within the current worksheet");
            }

            // If named range exists
            if (worksheet.NamedRangeExists(name))
            {
                throw new ArgumentException($"Name {name} already exists");
            }
            else
            {
                worksheet.Names.Add(name, worksheet.Range[$"{range}"]);
            }
        }

        /// <summary>
        /// Renames a range on the worksheet
        /// </summary>
        /// <param name="worksheet">Target worksheet</param>
        /// <param name="oldName">Old range name</param>
        /// <param name="newName">New range name</param>
        public static void RenameRange(this Excel.Worksheet worksheet, string oldName, string newName)
        {
            // Checks target range exists
            if (worksheet.NamedRangeExists(oldName))
            {
                // Checks if new name already exists
                if (!worksheet.NamedRangeExists(newName))
                {
                    worksheet.Range[$"{oldName}"].Name = newName;
                }
                else
                {
                    throw new ArgumentException($"New range name {newName} already exists. Please rename to something else.");
                }
            }
            else
            {
                throw new ArgumentException($"Range {oldName} does not exist on {worksheet.Name}");
            }
        }

        /// <summary>
        /// Gets name ranges from worksheet
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="name"></param>
        /// <returns></returns>
        public static Excel.Name GetNamedRange(this Excel.Worksheet worksheet, string name)
        {
            var names = worksheet.Names;
            foreach (Excel.Name item in worksheet.Names)
            {
                if (item.Name.Split('!').Last() == name)
                {
                    return item;
                }
            }

            throw new ArgumentException($"Name {name} does not exist");
        }

        /// <summary>
        /// Will clear the range of cells including formats
        /// </summary>
        /// <param name="worksheet">Target worksheet</param>
        /// <param name="name">Name of range</param>
        public static void ClearNamedRange(this Excel.Worksheet worksheet, string name)
        {
            if (worksheet.NamedRangeExists(name))
            {
                worksheet.Range[name].Clear();
            }
        }

        /// <summary>
        /// Deletes the worksheet's range contents 
        /// </summary>
        /// <param name="worksheet">Target worksheet</param>
        /// <param name="range">Range. Can either be name or range of cells. If not the name, declare next parameter as false</param>
        /// <param name="isNamedRange">Flags if range param is the name of the range</param>
        public static void DeleteRangeContents(this Excel.Worksheet worksheet, string range, bool isNamedRange = true)
        {
            if (isNamedRange)
            {
                if (worksheet.NamedRangeExists(range))
                {
                    worksheet.Range[$"{range}"].Cells.ClearContents();
                }
                else
                {
                    throw new ArgumentException($"Range, [{range}], does not exist");
                }
            }
            else
            {
                if (worksheet.IsRange(range))
                {
                    worksheet.Range[$"{range}"].Cells.ClearContents();
                }
                else
                {
                    throw new ArgumentException($"Range, [{range}], is not a valid range");
                }
            }
        }

        /// <summary>
        /// Removes named range from worksheet scope
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="name"></param>
        public static void RemoveNamedRange(this Excel.Worksheet worksheet, string name)
        {
            worksheet.GetNamedRange(name).Delete();
        }

        /// <summary>
        /// Returns the range of data that exists in the column. Range will start at the parameter cell and end before the first empty row it encounters
        /// </summary>
        /// <param name="worksheet">Target worksheet</param>
        /// <param name="range">Singular cell in the desired column</param>
        /// <returns>Excel range</returns>
        public static Excel.Range GetColumnRange(this Excel.Worksheet worksheet, string range)
        {
            if (worksheet.IsRange(range))
            {
                if (worksheet.Range[range].IsSingularCell())
                {
                    int rangeColumnNumber = worksheet.Range[range].Column;
                    int lastRow = worksheet.Range[range].Row;

                    while (!string.IsNullOrEmpty(worksheet.Range[$"{WorksheetHelper.GetColumnName(rangeColumnNumber)}" + $"{lastRow}"].Value))
                    {
                        lastRow++;
                    }

                    return worksheet.Range[$"{WorksheetHelper.GetColumnName(rangeColumnNumber)}" + $"{lastRow - 1}"];
                }
                else
                {
                    throw new Exception("Please ensure only one cell is referenced to indicate the range");
                }
            }
            else
            {
                throw new Exception($"The range, {range}, does not exist on the current worksheet, {worksheet.Name}");
            }
        }

        /// <summary>
        /// Returns the range of data that exists in the row. Range will start at the parameter cell and end before the first empty column it encounters
        /// </summary>
        /// <param name="worksheet">Target worksheet</param>
        /// <param name="range">Singular cell in the desired column</param>
        /// <returns>Excel range</returns>
        public static Excel.Range GetRowRange(this Excel.Worksheet worksheet, string range)
        {
            if (worksheet.IsRange(range))
            {
                if (worksheet.Range[range].IsSingularCell())
                {
                    int rangeRowNumber = worksheet.Range[range].Row;
                    int lastColumn = worksheet.Range[range].Column;

                    while (!string.IsNullOrEmpty(worksheet.Range[$"{WorksheetHelper.GetColumnName(lastColumn)}" + $"{rangeRowNumber}"].Value))
                    {
                        lastColumn++;
                    }

                    return worksheet.Range[$"{WorksheetHelper.GetColumnName(lastColumn - 1)}" + $"{rangeRowNumber}"];
                }
                else
                {
                    throw new Exception("Please ensure only one cell is referenced to indicate the range");
                }
            }
            else
            {
                throw new Exception($"The range, {range}, does not exist on the current worksheet, {worksheet.Name}");
            }
        }

        public static Excel.Range GetDataSetRange(this Excel.Worksheet worksheet, string range)
        {
            if (worksheet.IsRange(range))
            {
                if (worksheet.Range[range].IsSingularCell())
                {
                    int rangeRowNumber = worksheet.Range[range].Row;
                    int lastColumn = worksheet.Range[range].Column;

                    //Iterate through columns until blank cell is hit
                    while (!string.IsNullOrEmpty(worksheet.Range[$"{WorksheetHelper.GetColumnName(lastColumn)}" + $"{rangeRowNumber}"].Value))
                    {
                        lastColumn++;
                    }

                    int rangeColumnNumber = worksheet.Range[$"{WorksheetHelper.GetColumnName(lastColumn - 1)}" + $"{rangeRowNumber}"].Column;
                    int lastRow = worksheet.Range[$"{WorksheetHelper.GetColumnName(lastColumn - 1)}" + $"{rangeRowNumber}"].Row;

                    // Iterate through rows until blank cell is hit
                    while (!string.IsNullOrEmpty(worksheet.Range[$"{WorksheetHelper.GetColumnName(rangeColumnNumber)}" + $"{lastRow}"].Value))
                    {
                        lastRow++;
                    }

                    return worksheet.Range[$"{WorksheetHelper.GetColumnName(lastColumn - 1)}" + $"{lastRow - 1}"];
                }
                else
                {
                    throw new Exception("Please ensure only one cell is referenced to indicate the range");
                }
            }
            else
            {
                throw new Exception($"The range, {range}, does not exist on the current worksheet, {worksheet.Name}");
            }
        }

        public static Excel.Range GetHeaderCell(this Excel.Worksheet worksheet, string range, string headerText)
        {
            if (worksheet.IsRange(range))
            {
                if (worksheet.Range[range].IsSingularCell())
                {
                    int rangeRowNumber = worksheet.Range[range].Row;
                    int lastColumn = worksheet.Range[range].Column;

                    while (!string.IsNullOrEmpty(worksheet.Range[$"{WorksheetHelper.GetColumnName(lastColumn)}" + $"{rangeRowNumber}"].Value))
                    {
                        if (worksheet.Range[$"{WorksheetHelper.GetColumnName(lastColumn)}" + $"{rangeRowNumber}"].Value == headerText)
                        {
                            return worksheet.Range[$"{WorksheetHelper.GetColumnName(lastColumn)}" + $"{rangeRowNumber}"];
                        }
                        lastColumn++;
                    }

                    return null;
                }
                else
                {
                    throw new Exception("Please ensure only one cell is referenced to indicate the range");
                }
            }
            else
            {
                throw new Exception($"The range, {range}, does not exist on the current worksheet, {worksheet.Name}");
            }
        }
        #endregion

        #region Filters
        /// <summary>
        /// Removes any filtering applied to the current worksheet
        /// </summary>
        /// <param name="worksheet">Target worksheet</param>
        public static void ClearAutoFilters(this Excel.Worksheet worksheet)
        {
            if (worksheet.AutoFilter != null && worksheet.AutoFilterMode == true)
            {
                worksheet.AutoFilter.ShowAllData();
            }
        }
        #endregion
    }
}
