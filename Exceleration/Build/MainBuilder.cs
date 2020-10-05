using Exceleration.Commands;
using Exceleration.CoreHelpers;
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
        private WorkbookCommands _workbookCommands;
        private RangeCommands _rangeCommands;

        public MainBuilder(Excel.Worksheet worksheet) : base(worksheet)
        {
            _workbookCommands = new WorkbookCommands();
            _rangeCommands = new RangeCommands();
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
                    #region Workbook
                    case CommandType.Workbook:

                        switch (command)
                        {
                            case WorkbookCommands.AddSheet:
                                _workbookCommands.AddSheetCommand(workbook, value);                                                               
                                break;

                            case WorkbookCommands.CopySheet:
                                referenceType = OptionHelper.GetReferenceEnumFromString(reference);
                                _workbookCommands.CopySheetCommand(workbook, value, name, referenceType);
                                break;

                            case WorkbookCommands.DeleteSheet:
                                referenceType = OptionHelper.GetReferenceEnumFromString(reference);
                                _workbookCommands.DeleteSheetCommand(workbook, value);
                                break;

                            case WorkbookCommands.MoveSheet:

                                positional = OptionHelper.GetPositionalEnumFromString(option);
                                referenceType = OptionHelper.GetReferenceEnumFromString(reference);
                                _workbookCommands.MoveWorksheetCommand(workbook, value, positional, name, referenceType);
                                break;

                            case WorkbookCommands.TargetSheet:

                                referenceType = OptionHelper.GetReferenceEnumFromString(reference);
                                _workbookCommands.TargetSheetCommand(workbook, value, referenceType);
                                break;

                            case WorkbookCommands.RenameSheet:
                                referenceType = OptionHelper.GetReferenceEnumFromString(reference);
                                _workbookCommands.RenameSheetCommand(workbook, value, name, referenceType);
                                break;
                        }

                        break;

                    #endregion

                    #region Worksheet
                    case CommandType.Worksheet:

                        break;

                    #endregion

                    #region Range
                    case CommandType.Range:
                        switch (command)
                        {
                            case RangeCommands.AddNamedRange:
                                switch(option)
                                {
                                    case RangeOptions.WorkbookScope:
                                        _rangeCommands.AddWorkbookNamedRange(workbook, name, value);
                                        break;
                                    case RangeOptions.WorksheetScope:
                                        _rangeCommands.AddWorksheetNamedRange(workbook.ActiveSheet, name, value);
                                        break;
                                    default:
                                        _rangeCommands.AddWorkbookNamedRange(workbook, name, value);
                                        break;
                                }
                                
                                break;

                            case RangeCommands.DeleteRangeContents:
                                _rangeCommands.DeleteRangeContentsCommand(workbook.ActiveSheet, value, referenceType);
                                break;

                            case RangeCommands.RemoveNamedRange:
                                _rangeCommands.RemoveNamedRangeCommand(workbook.ActiveSheet, value);
                                break;

                            case RangeCommands.RenameRange:
                                switch (option)
                                {
                                    case RangeOptions.WorkbookScope:
                                        _rangeCommands.RenameWorkbookRange(workbook, value, name);
                                        break;
                                    case RangeOptions.WorksheetScope:
                                        _rangeCommands.RenameWorksheetRange(workbook.ActiveSheet, value, name);
                                        break;
                                    default:
                                        _rangeCommands.RenameWorkbookRange(workbook, value, name);
                                        break;
                                }

                                break;

                            case RangeCommands.SetNamedRange:
                                _rangeCommands.SetNamedRangeCommand(workbook.ActiveSheet, name, value);
                                break;
                        }

                        break;
                        #endregion
                }

                i++;
            }
           
            return new List<string>();
        }
    }
}
