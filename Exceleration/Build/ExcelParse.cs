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

        protected string GetString(int row, int column)
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

        protected bool GetBoolean(int row, int column)
        {
            var value = GetString(row, column);

            return value.ToUpper() == "TRUE" || value.ToUpper() == "YES";
        }

        protected double GetDouble(int row, int column)
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
        /// Convert value in Excel cell to an integer value
        /// </summary>
        /// <param name="row">Row number</param>
        /// <param name="column">Column number</param>
        /// <returns></returns>
        protected int GetInt(int row, int column)
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
        /// Checks if commands have matching end if commands
        /// </summary>
        /// <param name="row">Starting row of if statement</param>
        /// <param name="commandColumn">Column to loop through and find matching end if command</param>
        protected void ValidateIf(int row, int commandColumn)
        {
            var rowStart = row;
            var startRepeat = 0;
            var endRepeat = 0;

            while (!string.IsNullOrEmpty(GetString(row, commandColumn)))
            {
                var command = GetString(row, commandColumn).ToUpper();

                if (command == CodeCommands.If)
                {
                    startRepeat++;
                }

                if (command == CodeCommands.EndIf)
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
                throw new Exception($"If on line {rowStart} on worksheet {_worksheet.Name} does not have a matching end if");
            }
        }

        /// <summary>
        /// Returns the row number where the matching end if is located
        /// </summary>
        /// <param name="row">Row number</param>
        /// <param name="commandColumn">Column to loop through and find matching end if command</param>
        /// <returns></returns>
        protected int GetEndIfRow(int row, int commandColumn)
        {
            var ifcount = 0;
            var endIfCount = 0;

            var endIfLocations = new List<int>();

            while (!string.IsNullOrEmpty(GetString(row, commandColumn)))
            {
                var command = GetString(row, commandColumn).ToUpper();

                if (command == CodeCommands.If)
                {
                    ifcount++;
                }

                if (command == CodeCommands.EndIf)
                {
                    endIfCount++;
                    endIfLocations.Add(row);
                }

                if (ifcount == endIfCount)
                {
                    break;
                }

                row++;
            }

            return endIfLocations.Max();
        }

        protected int GetEndRepeatRow(int row, int commandColumn)
        {
            var rowStart = row;
            var repeatCount = 0;
            var endRepeatCount = 0;

            var endRepeatLocation = new List<int>();

            while (!string.IsNullOrEmpty(GetString(row, commandColumn)))
            {
                var command = GetString(row, commandColumn).ToUpper();

                if (command == CodeCommands.Repeat)
                {
                    repeatCount++;
                }

                if (command == CodeCommands.EndRepeat)
                {
                    endRepeatCount++;
                    endRepeatLocation.Add(row);
                }

                if (repeatCount == endRepeatCount)
                {
                    break;
                }

                row++;
            }

            return endRepeatLocation.Max();
        }

        protected void ValidateRepeat(int row, int commandColumn)
        {
            var rowStart = row;
            var startRepeat = 0;
            var endRepeat = 0;

            while (!string.IsNullOrEmpty(GetString(row, commandColumn)))
            {
                var command = GetString(row, commandColumn).ToUpper();

                if (command == CodeCommands.Repeat)
                {
                    startRepeat++;
                }

                if (command == CodeCommands.EndRepeat)
                {
                    endRepeat++;
                }

                row++;
            }

            if (startRepeat != endRepeat)
            {
                throw new Exception($"Start repeat on line {rowStart} on worksheet {_worksheet.Name} does not have a matching end repeat");
            }

            if (startRepeat > 1 || endRepeat > 1)
            {
                throw new Exception("Nested repeats are not supported");
            }
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
