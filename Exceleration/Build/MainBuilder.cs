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
using Exceleration.Helpers.Extensions;

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

        public List<string> Run(bool startMethod, string rangeName = "")
        {
            Excel.Range startCell = null;
            Excel.Workbook workbook = Globals.ThisAddIn.Application.ActiveWorkbook;

            try
            {
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
            }
            catch
            {
                throw new ArgumentException("Issue finding starting code cell. Please check that start cells match their respective sheet names.");
            }
            

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
                string option;
                string reference;
                string name;
                string value;
                ReferenceEnum referenceType = ReferenceEnum.ByName;
                PositionalEnum positional = PositionalEnum.AtEnd;

                //var command = GetString(i, commandColumn).ToUpper();
                //var option = GetString(i, optionsColumn).ToUpper();
                //var reference = GetString(i, referenceColumn).ToUpper();
                //var name = GetString(i, nameColumn);
                //var value = GetString(i, valueColumn);


                // Ignores comment lines
                if (command == CodeCommands.Comment)
                {
                    i++;
                    continue;
                }
                else
                {
                    command = GetString(i, commandColumn).ToUpper();
                    option = GetString(i, optionsColumn).ToUpper();
                    reference = GetString(i, referenceColumn).ToUpper();
                    name = GetString(i, nameColumn);
                    value = GetString(i, valueColumn);
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
                                switch (option)
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

                                referenceType = OptionHelper.GetReferenceEnumFromString(reference);
                                switch (option)
                                {
                                    case RangeOptions.WorkbookScope:
                                        _rangeCommands.DeleteWorkbookRangeContents(workbook, value, referenceType);
                                        break;
                                    case RangeOptions.WorksheetScope:
                                        _rangeCommands.DeleteWorksheetRangeContents(workbook.ActiveSheet, value, referenceType);
                                        break;
                                    default:
                                        _rangeCommands.DeleteWorkbookRangeContents(workbook, value, referenceType);
                                        break;
                                }

                                break;

                            case RangeCommands.RemoveNamedRange:

                                switch(option)
                                {
                                    case RangeOptions.WorkbookScope:
                                        _rangeCommands.RemoveWorkbookNamedRange(workbook, value);
                                        break;
                                    case RangeOptions.WorksheetScope:
                                        _rangeCommands.RemoveWorksheetNamedRange(workbook.ActiveSheet, value);
                                        break;
                                    default:
                                        _rangeCommands.RemoveWorkbookNamedRange(workbook, value);
                                        break;
                                }

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
                                _rangeCommands.SetNamedRangeCommand(workbook, name, value);
                                break;
                        }

                        break;
                    #endregion

                    case CommandType.Code:

                        switch (command)
                        {
                            case CodeCommands.Sub:
                                MainBuilder runblock = null;
                                var subName = GetString(i, valueColumn);
                                string worksheetName;

                                // Allows for calling sub-routines on different worksheets
                                if (subName.Contains("!"))
                                {
                                    var subSplit = subName.Split('!');

                                    worksheetName = subSplit[0];
                                    subName = subSplit[1];

                                    runblock = new MainBuilder(workbook.GetWorksheet(worksheetName));
                                }
                                else
                                {
                                    runblock = new MainBuilder(_worksheet);
                                }

                                runblock.Run(false, subName);
                                break;
                        }

                        break;
                }

                i++;
            }
           
            return new List<string>();
        }
    }
}
