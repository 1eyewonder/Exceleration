using Exceleration.Commands;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace Exceleration.Build
{
    public class ExcelParse : IDisposable
    {
        protected Excel.Worksheet _worksheet;

        public List<string> Logs { get; set; }


        public ExcelParse(Excel.Worksheet worksheet)
        {
            _worksheet = worksheet;
            Logs = new List<string>();
        }

        #region Value Helpers 

        protected string _getString(int row, int column)
        {
            var range = _worksheet.Cells[row, column];

            var value = "";

            if (range.Value == null)
            {
                value = "";
            }
            else
            {
                value = range.Value.ToString();
            }

            if (range != null)
            {
                Marshal.ReleaseComObject(range);
            }

            return value;
        }

        /// <summary>
        /// Converts value in Excel cell to a list of string base on comma delimination
        /// </summary>
        /// <param name="row">Row number</param>
        /// <param name="column">Column number</param>
        /// <returns></returns>
        protected string[] _getArrayString(int row, int column)
        {
            var range = _worksheet.Cells[row, column];
            string[] theList = null;

            if (range.Value != null)
            {
                string someValue = range.Value.ToString();
                theList = someValue.Split(',');
            }

            return theList;
        }

        protected bool _getBoolean(int row, int column)
        {
            var value = _getString(row, column);

            return value.ToUpper() == "TRUE" || value.ToUpper() == "YES";
        }

        protected double _getDouble(int row, int column)
        {
            var range = _worksheet.Cells[row, column];

            var value = 0.00;

            if (range.Value == null)
            {
                value = 0.00;
            }
            else
            {
                if (!double.TryParse(range.Value.ToString(), out value))
                {
                    value = 0.00;
                }
            }

            if (range != null)
            {
                Marshal.ReleaseComObject(range);
            }

            return value;
        }

        /// <summary>
        /// Convert value in Excel cell to an integer value. Returns 0 if empty
        /// </summary>
        /// <param name="row">Row number</param>
        /// <param name="column">Column number</param>
        /// <returns></returns>
        protected int _getInt(int row, int column)
        {
            var range = _worksheet.Cells[row, column];

            var value = 0;

            if (range.Value == null)
            {
                value = 0;
            }
            else
            {
                if (!int.TryParse(range.Value.ToString(), out value))
                {
                    throw new Exception($"Value entered ({range.Value}) on {_worksheet.Name}, line {row} is not an integer. Please enter a valid integer.");
                }
            }

            if (range != null)
            {
                Marshal.ReleaseComObject(range);
            }

            return value;
        }

        public void SetValue(int row, int column, string value)
        {
            var range = _worksheet.Cells[row, column];

            range.Value = value;

            if (range != null)
            {
                Marshal.ReleaseComObject(range);
            }
        }

        #endregion

        #region Generic Command Validation

        /// <summary>
        /// Checks if code command has matching keyword to end its loop
        /// </summary>
        /// <param name="row">Starting row of the code command</param>
        /// <param name="commandColumn">Column to loop through and find matching end commands</param>
        /// <param name="startCommand">Start command name</param>
        /// <param name="endCommand">End command name</param>
        /// Example start and end commands can be found in the CodeCommands class
        protected void _validateCodeCommand(int row, int commandColumn, string startCommand, string endCommand)
        {
            var rowStart = row;
            var startRepeat = 0;
            var endRepeat = 0;

            while (!string.IsNullOrEmpty(_getString(row, commandColumn)))
            {
                var command = _getString(row, commandColumn).ToUpper();

                if (command == startCommand)
                {
                    startRepeat++;
                }

                if (command == endCommand)
                {
                    endRepeat++;
                }

                if (startRepeat == endRepeat)
                {
                    break;
                }

                row++;
            }

            if (startRepeat != endRepeat)
            {
                throw new Exception($"{startCommand} on line {rowStart} on worksheet {_worksheet.Name} does not have a matching {endCommand}");
            }
        }

        /// <summary>
        /// Returns the row number where the match end command is located
        /// </summary>
        /// <param name="row">Starting row of the code command</param>
        /// <param name="commandColumn">Column to loop through and find matching end commands</param>
        /// <param name="startCommand">Start command name</param>
        /// <param name="endCommand">End command name</param>
        /// <returns></returns>
        protected int _getEndCommandRow(int row, int commandColumn, string startCommand, string endCommand)
        {
            var startCommandCount = 0;
            var endCommandCount = 0;

            var endCommandLocations = new List<int>();

            while (!string.IsNullOrEmpty(_getString(row, commandColumn)))
            {
                var command = _getString(row, commandColumn).ToUpper();

                if (command == startCommand)
                {
                    startCommandCount++;
                }

                if (command == endCommand)
                {
                    endCommandCount++;
                    endCommandLocations.Add(row);
                }

                if (startCommandCount == endCommandCount)
                {
                    break;
                }

                row++;
            }

            return endCommandLocations.Max();
        }

        #endregion

        public void Dispose()
        {
            if (_worksheet != null)
            {
                Marshal.ReleaseComObject(_worksheet);
            }
        }
    }
}
