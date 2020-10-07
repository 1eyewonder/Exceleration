using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace Exceleration.Helpers.Extensions
{
    public static class ExcelRangeExtensions
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
