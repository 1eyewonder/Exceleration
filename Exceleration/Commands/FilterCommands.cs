using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;
using Exceleration.Helpers.Extensions;

namespace Exceleration.Commands
{
    public class FilterCommands
    {
        public const string AddDataFilter = "ADD DATA FILTER";
        public const string ClearDataFilters = "CLEAR DATA FILTERS";
        public const string DeleteDataFilters = "DELETE DATA FILTER";

        public void AddDataFilterCommand(Excel.Worksheet worksheet, string range, int columnNumber, Excel.XlAutoFilterOperator filterEnum, string[] criteria = null, string criteria2 = null)
        {
            if (!worksheet.IsRange(range))
            {
                throw new Exception($"The range, {range}, does not exist on the current worksheet, {worksheet.Name}");
            }

            // ExcelParse class returns 0 as integer value for blank cells and column indices start at 1 for Excel
            if (columnNumber == 0)
            {
                columnNumber = 1;
            }

            if (criteria != null && criteria.Length > 0)
            {
                if (string.IsNullOrEmpty(criteria2))
                {
                    worksheet.Range[range].AddAutoFilter(columnNumber, filterEnum, criteria);
                }
                else
                {
                    worksheet.Range[range].AddAutoFilter(columnNumber, filterEnum, criteria, criteria2);
                }
            }
            else
            {
                worksheet.Range[range].AddAutoFilter(columnNumber, filterEnum);
            } 
        }

        public void ClearDataFiltersCommand(Excel.Worksheet worksheet)
        {
            worksheet.ClearAutoFilters();
        }

        public void DeleteDataFiltersCommand(Excel.Worksheet worksheet)
        {
            worksheet.AutoFilterMode = false;
        }
    }
}
