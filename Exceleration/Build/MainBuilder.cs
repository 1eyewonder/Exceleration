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
        private readonly WorkbookCommands _workbookCommands;
        private readonly WorksheetCommands _worksheetCommands;
        private readonly RangeCommands _rangeCommands;
        private readonly FilterCommands _filterCommands;
        private readonly DataCommands _dataCommands;
        private bool _inRepeat = false;
        private int _repeatStart = 0;
        private int _repeatEnd = 0;
        private int _repeatCount = 0;
        private int _repeatIndex = 0;

        public MainBuilder(Excel.Worksheet worksheet) : base(worksheet)
        {
            _workbookCommands = new WorkbookCommands();
            _rangeCommands = new RangeCommands();
            _worksheetCommands = new WorksheetCommands();
            _filterCommands = new FilterCommands();
            _dataCommands = new DataCommands();
        }

        /// <summary>
        /// Runs code
        /// </summary>
        /// <param name="startMethod">Flags method as first called method</param>
        /// <param name="rangeName">Named range indicating code blocks starting position</param>
        /// <returns></returns>
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
            var targetColumn = nameColumn + 1;
            var auxillaryColumn = targetColumn + 1;

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
                string target;
                string auxillary;

                ReferenceEnum referenceType = ReferenceEnum.ByName;
                PositionalEnum positional = PositionalEnum.AtEnd;

                // Ignores comment lines
                if (command == CodeCommands.Comment)
                {
                    i++;
                    continue;
                }
                else
                {
                    option = GetString(i, optionsColumn).ToUpper();
                    reference = GetString(i, referenceColumn).ToUpper();
                    name = GetString(i, nameColumn);
                    target = GetString(i, targetColumn);
                    auxillary = GetString(i, auxillaryColumn);
                }
               
                switch (commandType)
                {
                    #region Workbook
                    case CommandType.Workbook:

                        switch (command)
                        {
                            case WorkbookCommands.AddSheet:
                                _workbookCommands.AddSheetCommand(workbook, target);                                                               
                                break;

                            case WorkbookCommands.CopySheet:
                                referenceType = OptionHelper.GetReferenceEnumFromString(reference);
                                _workbookCommands.CopySheetCommand(workbook, target, name, referenceType);
                                break;

                            case WorkbookCommands.DeleteSheet:
                                referenceType = OptionHelper.GetReferenceEnumFromString(reference);
                                _workbookCommands.DeleteSheetCommand(workbook, target, referenceType);
                                break;

                            case WorkbookCommands.MoveSheet:

                                positional = OptionHelper.GetPositionalEnumFromString(option);
                                referenceType = OptionHelper.GetReferenceEnumFromString(reference);
                                _workbookCommands.MoveWorksheetCommand(workbook, target, positional, name, referenceType);
                                break;

                            case WorkbookCommands.TargetSheet:

                                referenceType = OptionHelper.GetReferenceEnumFromString(reference);
                                _workbookCommands.TargetSheetCommand(workbook, target, referenceType);
                                break;

                            case WorkbookCommands.RenameSheet:
                                referenceType = OptionHelper.GetReferenceEnumFromString(reference);
                                _workbookCommands.RenameSheetCommand(workbook, target, name, referenceType);
                                break;
                        }

                        break;

                    #endregion

                    #region Worksheet
                    case CommandType.Worksheet:
                        switch(command)
                        {
                            case WorksheetCommands.AddColumn:
                                _worksheetCommands.AddColumnCommand(workbook.ActiveSheet, target);
                                break;
                            case WorksheetCommands.AddRow:
                                _worksheetCommands.AddRowCommand(workbook.ActiveSheet, target);
                                break;
                            case WorksheetCommands.MoveColumn:
                                _worksheetCommands.MoveColumnCommand(workbook.ActiveSheet, target, name);
                                break;
                            case WorksheetCommands.MoveRow:
                                _worksheetCommands.MoveRowCommand(workbook.ActiveSheet, target, name);
                                break;
                            case WorksheetCommands.DeleteColumn:
                                _worksheetCommands.DeleteColumnCommand(workbook.ActiveSheet, target);
                                break;
                            case WorksheetCommands.DeleteRow:
                                _worksheetCommands.DeleteRowCommand(workbook.ActiveSheet, target);
                                break;
                        }

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
                                        _rangeCommands.AddWorkbookNamedRange(workbook, name, target);
                                        break;
                                    case RangeOptions.WorksheetScope:
                                        _rangeCommands.AddWorksheetNamedRange(workbook.ActiveSheet, name, target);
                                        break;
                                    default:
                                        _rangeCommands.AddWorkbookNamedRange(workbook, name, target);
                                        break;
                                }
                                break;

                            case RangeCommands.DeleteRangeContents:

                                referenceType = OptionHelper.GetReferenceEnumFromString(reference);
                                switch (option)
                                {
                                    case RangeOptions.WorkbookScope:
                                        _rangeCommands.DeleteWorkbookRangeContents(workbook, target, referenceType);
                                        break;
                                    case RangeOptions.WorksheetScope:
                                        _rangeCommands.DeleteWorksheetRangeContents(workbook.ActiveSheet, target, referenceType);
                                        break;
                                    default:
                                        _rangeCommands.DeleteWorkbookRangeContents(workbook, target, referenceType);
                                        break;
                                }

                                break;

                            case RangeCommands.RemoveNamedRange:

                                switch(option)
                                {
                                    case RangeOptions.WorkbookScope:
                                        _rangeCommands.RemoveWorkbookNamedRange(workbook, target);
                                        break;
                                    case RangeOptions.WorksheetScope:
                                        _rangeCommands.RemoveWorksheetNamedRange(workbook.ActiveSheet, target);
                                        break;
                                    default:
                                        _rangeCommands.RemoveWorkbookNamedRange(workbook, target);
                                        break;
                                }

                                break;

                            case RangeCommands.RenameRange:
                                switch (option)
                                {
                                    case RangeOptions.WorkbookScope:
                                        _rangeCommands.RenameWorkbookRange(workbook, target, name);
                                        break;
                                    case RangeOptions.WorksheetScope:
                                        _rangeCommands.RenameWorksheetRange(workbook.ActiveSheet, target, name);
                                        break;
                                    default:
                                        _rangeCommands.RenameWorkbookRange(workbook, target, name);
                                        break;
                                }

                                break;

                            case RangeCommands.SetNamedRange:
                                _rangeCommands.SetNamedRangeCommand(workbook, name, target);
                                break;
                        }

                        break;
                    #endregion

                    #region Code
                    case CommandType.Code:

                        switch (command)
                        {
                            case CodeCommands.Sub:
                                MainBuilder runblock = null;
                                var subName = GetString(i, targetColumn);
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

                            case CodeCommands.If:
                                ValidateIf(i, commandColumn);

                                var booleanValue = GetBoolean(i, targetColumn);

                                // Skips to end if command line if false
                                if (!booleanValue)
                                {
                                    i = GetEndIfRow(i, commandColumn);
                                }

                                break;

                            case CodeCommands.Stop:
                                throw new Exception("Program stopped");

                            case CodeCommands.Repeat:
                                ValidateRepeat(i, commandColumn);
                                _repeatStart = i;
                                _repeatEnd = GetEndRepeatRow(i, commandColumn);
                                _repeatCount = GetInt(i, targetColumn);
                                _repeatIndex = 1;
                                SetValue(i, targetColumn + 1, _repeatIndex.ToString());
                                _inRepeat = true;
                                break;
                        }

                        break;
                    #endregion

                    #region Filter
                    case CommandType.Filter:

                        switch(command)
                        {
                            case FilterCommands.AddDataFilter:
                                var referenceInt = GetInt(i, referenceColumn);                              
                                var criteriaList = GetArrayString(i, nameColumn);
                                auxillary = GetString(i, auxillaryColumn);

                                var filterType = OptionHelper.GetExcelAutoFilterOperatorFromString(option);                             
                                 _filterCommands.AddDataFilterCommand(workbook.ActiveSheet, target, referenceInt, filterType, criteriaList, auxillary);

                                break;

                            case FilterCommands.ClearDataFilters:
                                _filterCommands.ClearDataFiltersCommand(workbook.ActiveSheet);
                                break;

                            case FilterCommands.DeleteDataFilters:
                                _filterCommands.DeleteDataFiltersCommand(workbook.ActiveSheet);
                                break;
                            case "TEST":
                                
                                break;
                        }

                        break;
                    #endregion

                    #region Data
                    case CommandType.Data:

                        switch(command)
                        {
                            case DataCommands.SetValue:
                                break;

                            case DataCommands.FindAndReplace:
                                bool matchCase = GetBoolean(i, referenceColumn);
                                _dataCommands.FindAndReplaceCommand(workbook.ActiveSheet, target, name, auxillary, option, matchCase);
                                break;
                        }

                        break;
                    #endregion
                }

                i++;

                if (_inRepeat)
                {
                    if (i == _repeatEnd)
                    {
                        if (_repeatIndex == _repeatCount)
                        {
                            i = _repeatEnd + 1;
                            _repeatStart = 0;
                            _repeatEnd = 0;
                            _repeatCount = 0;
                            _repeatIndex = 0;
                            _inRepeat = false;
                        }
                        else
                        {
                            i = _repeatStart + 1;
                            _repeatIndex++;
                            SetValue(_repeatStart, targetColumn+1, _repeatIndex.ToString());
                        }
                    }
                }
            }
           
            return new List<string>();
        }
    }
}
