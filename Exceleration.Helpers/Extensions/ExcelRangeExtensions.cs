using System;
using System.Collections.Generic;
using System.Data;
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
            range.Validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertInformation, Excel.XlFormatConditionOperator.xlBetween, $"={rangeName}", Type.Missing);
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

            foreach (Excel.Range r in range)
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

        /// <summary>
        /// Returns true if range consists of one cell
        /// </summary>
        /// <param name="range"></param>
        /// <returns></returns>
        public static bool IsSingularCell(this Excel.Range range)
        {
            if (range.Cells.Count == 1)
            {
                return true;
            }
            else return false;
        }

        /// <summary>
        /// Adds a data filter to a range
        /// </summary>
        /// <param name="range">Target range, including the header row</param>
        /// <param name="columnNumber">Column to apply filter to, index starts at 1 and relates to the first column on the left</param>
        /// <param name="filterEnum">Excel enum how data will be filtered</param>
        /// <param name="criteria">First criteria for filtering</param>
        /// <param name="criteria2">Second criteria for filtering</param>
        public static void AddAutoFilter(this Excel.Range range, int columnNumber = 1, Excel.XlAutoFilterOperator filterEnum = Excel.XlAutoFilterOperator.xlAnd, string[] criteria = null, string criteria2 = null)
        {
            int columnCount = range.Columns.Count;

            if (columnNumber < 1 || columnNumber > columnCount)
            {
                throw new Exception($"The range selected, {range}, only has {columnCount} columns. Please enter a value greater than or equal to 1 and less than or equal to {columnCount}");
            }

            string singularCriteria = "";
            if (criteria.Length == 1)
            {
                singularCriteria = criteria.First();
            }

            // Adds data filter to range without filtering data
            if (columnNumber == 1 && filterEnum == Excel.XlAutoFilterOperator.xlAnd && criteria == null && criteria2 == null)
            {
                range.AutoFilter(1);
            }

            // If filter can be applied to a column without criteria
            else if (criteria == null && criteria2 == null)
            {
                range.AutoFilter(columnNumber, Type.Missing, filterEnum);
            }

            // If there is one criteria
            else if (criteria != null && criteria2 == null)
            {
                if (criteria.Length == 1)
                {
                    range.AutoFilter(columnNumber, singularCriteria, filterEnum);                   
                }
                else
                {
                    range.AutoFilter(columnNumber, criteria, filterEnum);
                }               
            }

            // If there are two criteria
            else
            {
                if (criteria.Length == 1)
                {
                    range.AutoFilter(columnNumber, singularCriteria, filterEnum, criteria2);
                }
                else
                {
                    range.AutoFilter(columnNumber, criteria, filterEnum, criteria2);
                }                
            }            
        }

        /// <summary>
        /// Find and replace text within a range
        /// </summary>
        /// <param name="range">Target range</param>
        /// <param name="oldText">Text to find</param>
        /// <param name="newText">To to replace with</param>
        /// <param name="whole">True if whole cell string is to be replaced</param>
        /// <param name="matchCase">True if oldText case matters in search</param>
        /// Per Microsoft documentation, the active selection/active cell is not affected
        public static void FindAndReplace(this Excel.Range range, string oldText, string newText, bool whole = false, bool matchCase = false)
        {
            if (whole)
            {
                if (matchCase)
                {
                    range.Replace(oldText, newText, Excel.XlLookAt.xlWhole, Excel.XlSearchOrder.xlByRows, true);
                }
                else
                {
                    range.Replace(oldText, newText, Excel.XlLookAt.xlWhole, Excel.XlSearchOrder.xlByRows, false);
                }
            }
            else
            {
                if (matchCase)
                {
                    range.Replace(oldText, newText, Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, true);
                }
                else
                {
                    range.Replace(oldText, newText, Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, true);
                }
            }
        }

        public static DataTable ConvertToDataTable(this Excel.Range range, Dictionary<string, Type> propertyMap)
        {
            // Initialize data table and convert range to array
            DataTable dataTable = new DataTable();
            object[,] data = range.Value2;
            
            for (int columnCount = 1; columnCount <= range.Columns.Count; columnCount++)
            {
                // Create new column in data table
                var column = new DataColumn();
                string columnName = data[1, columnCount].ToString();

                // Checks what type is associate with the column property
                if (propertyMap.TryGetValue(columnName, out var output))
                {
                    column.DataType = output;
                    column.ColumnName = columnName;
                    dataTable.Columns.Add(column);
                }
                else
                {
                    throw new Exception();
                }                        

                for (int rowCount = 2; rowCount <= range.Rows.Count; rowCount ++)
                {                    
                    dynamic cellValue;
                    DataRow row;

                    // Converts value in array to type from property map
                    var item = data[rowCount, columnCount];
                    cellValue = Convert.ChangeType(item, output);                    

                    // Creates new row in data table if first item for the row
                    if (columnCount == 1)
                    {
                        row = dataTable.NewRow();
                        row[columnName] = cellValue;
                        dataTable.Rows.Add(row);
                    }
                    else
                    {
                        row = dataTable.Rows[rowCount-2];
                        row[columnName] = cellValue;
                    }
                }                
            }

            return dataTable;
        }
    }
}
