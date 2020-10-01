using Exceleration.Commands;
using Exceleration.Helpers.Enums;
using Exceleration.Options;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace Exceleration.Build
{   
    public class MainBuilder : ExcelParse
    {
        private SheetCommands _sheetCommands;

        public MainBuilder(Excel.Worksheet worksheet) : base(worksheet)
        {
            _sheetCommands = new SheetCommands();
        }

        public List<string> Run(bool startMethod, string rangeName = "", string rangeParameter = "")
        {
            Excel.Range startCell = null;
            Excel.Workbook workbook = Globals.ThisAddIn.Application.ActiveWorkbook;

            // Looks for range to start on
            if (string.IsNullOrEmpty(rangeName))
            {
                // Starting cell is named range that matches the worksheet the run command is being called on
                startCell = _worksheet.Range[_worksheet.Name + "Type"];
            }
            else
            {
                startCell = _worksheet.Range[rangeName];
            }

            if (startCell == null) throw new Exception("Could not find start range");

            // Sets starting integers for columns
            var typeColumn = startCell.Column;
            var commandColumn = typeColumn + 1;
            var optionsColumn = commandColumn + 1;
            var referenceColumn = optionsColumn + 1;
            var nameColumn = referenceColumn + 1;
            var valueColumn = nameColumn + 1;

            // Starts at cell below starting range for the sheet
            int i = startCell.Row + 1;

            // Will continue to run until there is an empty row found in the command type column
            while (!string.IsNullOrEmpty(GetString(i, typeColumn)))
            {
                var commandType = GetString(i, typeColumn).ToUpper();
                var command = GetString(i, commandColumn).ToUpper();
                var option = GetString(i, optionsColumn).ToUpper();
                var reference = GetString(i, referenceColumn).ToUpper();
                var name = GetString(i, nameColumn);
                var value = GetString(i, valueColumn);
                ReferenceEnum referenceType = ReferenceEnum.ByName;
                PositionalEnum positional = PositionalEnum.AtEnd;

                // Ignores comment lines
                if (commandType == CommandType.Comment)
                {
                    i++;
                    continue;
                }
               
                switch (commandType)
                {
                    #region Worksheet
                    case CommandType.Worksheet:

                        switch (command)
                        {
                            case SheetCommands.AddSheet:
                                _sheetCommands.AddSheetCommand(workbook, value);                                                               
                                break;

                            case SheetCommands.CopySheet:
                                _sheetCommands.CopySheetCommand(workbook, value, name);
                                break;

                            case SheetCommands.DeleteSheet:
                                _sheetCommands.DeleteSheetCommand(workbook, value);
                                break;

                            case SheetCommands.MoveSheet:
                                // Determines position desired by user
                                switch (option)
                                {
                                    case (SheetOptions.After):
                                        positional = PositionalEnum.After;
                                        break;
                                    case (SheetOptions.AtBeginning):
                                        positional = PositionalEnum.AtBeginning;
                                        break;
                                    case (SheetOptions.AtEnd):
                                        positional = PositionalEnum.AtEnd;
                                        break;
                                    case (SheetOptions.Before):
                                        positional = PositionalEnum.Before;
                                        break;
                                    default:
                                        positional = PositionalEnum.AtEnd;
                                        break;
                                }

                                // If relative sheet for positioning is required
                                if (positional != PositionalEnum.AtBeginning && positional != PositionalEnum.AtEnd)
                                {
                                    // Checks positional sheet reference is declared
                                    if (string.IsNullOrEmpty(reference) && string.IsNullOrEmpty(name))
                                    {
                                        break; // add log error
                                    }
                                }

                                // Checks if specific name is entered
                                if (string.IsNullOrEmpty(value))
                                {
                                    throw new ArgumentException("Must enter a desired worksheet name in the value column");
                                }

                                switch (reference)
                                {
                                    case (ReferenceOptions.ByName):
                                        referenceType = ReferenceEnum.ByName;
                                        break;

                                    case (ReferenceOptions.ByIndex):
                                        referenceType = ReferenceEnum.ByIndex;
                                        break;

                                    default:
                                        referenceType = ReferenceEnum.ByName;
                                        break;

                                }

                                _sheetCommands.MoveWorksheetCommand(workbook, value, positional, name, referenceType);
                                break;
                        }

                        break;

                    #endregion

                    #region Workbook
                    case CommandType.Workbook:

                        break;

                    #endregion
                }

                i++;
            }
           
            return new List<string>();
        }
    }
}
