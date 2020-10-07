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
    }
}
