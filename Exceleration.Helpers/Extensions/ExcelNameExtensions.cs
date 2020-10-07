using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace Exceleration.Helpers.Extensions
{
    public static class ExcelNameExtensions
    {
        /// <summary>
        /// Returns the name of an Excel.Name object without the sheet prefix, if any
        /// </summary>
        /// <param name="name">Excel.Name object</param>
        /// <returns></returns>
        public static string ShortName(this Excel.Name name)
        {
            if (name.Name.Contains('!'))
            {
                return name.Name.Split('!').Last();
            }
            else
            {
                return name.Name;
            }
        }
    }
}
