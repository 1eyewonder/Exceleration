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
    public static class RangeHelper
    {
        /// <summary>
        /// Adds a dropdown list to the given range
        /// </summary>
        /// <param name="range">Range where dropdown list is desired</param>
        /// <param name="rangeName">Name of range where options are located</param>
        public static void AddDropDownList(this Excel.Range range, string rangeName)
        {
            range.Validation.Delete();
            range.Validation.Add(XlDVType.xlValidateList, XlDVAlertStyle.xlValidAlertInformation, XlFormatConditionOperator.xlBetween, $"={rangeName}", Type.Missing);
            range.Validation.IgnoreBlank = false;
            range.Validation.InCellDropdown = true;
        }

        /// <summary>
        /// Checks if range exists in the given worksheet
        /// </summary>
        /// <param name="workSheet">Target worksheet</param>
        /// <param name="name">Name of range user is checking</param>
        /// <returns></returns>
        public static bool RangeExists(this Excel.Worksheet workSheet, string name)
        {
            return workSheet.NamedRanges().Exists(x => x == name);
        }

        /// <summary>
        /// Returns a list of all named ranges that exist on the target worksheet
        /// </summary>
        /// <param name="workSheet">Target worksheet</param>
        /// <returns></returns>
        public static List<string> NamedRanges(this Excel.Worksheet workSheet)
        {
            var value = new List<string>();

            foreach (Excel.Name n in workSheet.Application.ActiveWorkbook.Names)
            {
                value.Add(n.Name);
            }

            return value;
        }

        /// <summary>
        /// Creates a named range at the target location
        /// </summary>
        /// <param name="worksheet">Target worksheet</param>
        /// <param name="name">Name of range being created</param>
        /// <param name="range">Target range</param>
        /// <returns></returns>
        public static Excel.Range CreateNamedRange(this Excel.Worksheet worksheet, string name, string range)
        {
            Excel.Range namedRange = null;

            // If named range exists
            if (worksheet.RangeExists(name))
            {
                // Names range
                namedRange = worksheet.Range[$"{name}"]; // issue with this code

                // Assigns cell range to target
                namedRange = worksheet.Range[range];
            }
            else
            {
                // Assigns cell range to target
                namedRange = worksheet.Range[range];

                // Names range
                namedRange.Name = name;
            }

            return namedRange;
        }

        /// <summary>
        /// Removes a range from a worksheet's list of named ranges
        /// </summary>
        /// <param name="worksheet">Target worksheet</param>
        /// <param name="name">Name of range</param>
        public static void ClearNamedRange(this Excel.Worksheet worksheet, string name)
        {
            if (worksheet.RangeExists(name))
            {
                worksheet.Range[name].Clear();
            }
        }

        /// <summary>
        /// Returns a list of valid file paths in a target range
        /// </summary>
        /// <param name="range">Target range</param>
        /// <returns></returns>
        public static List<string> GetFilePaths(this Excel.Range range)
        {
            var tempList = new List<string>();

            foreach (Range r in range)
            {
                if (r.Value != null)
                {
                    var text = r.Value.ToString();

                    if (!string.IsNullOrEmpty(text))
                    {
                        if (FileHelper.IsValidPath(text))
                        {
                            tempList.Add(text);
                        }
                    }
                }
            }
            return tempList;
        }
    }
}
