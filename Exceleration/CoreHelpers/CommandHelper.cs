﻿using Exceleration.Options;
using Exceleration.DTOS;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Exceleration.Commands;

namespace Exceleration.CoreHelpers
{
    public static class CommandHelper
    {
        public static List<CommandItem> GetWorkbookCommands()
        {
            return new List<CommandItem>
            {
                new CommandItem(CommandType.Workbook, WorkbookCommands.AddSheet,"Not needed for this command","Not needed for this command","Not needed for this command", "Name of new worksheet being added. Use of '=' or '!' in sheet names will potentially cause issues.","Not needed for this command"),
                new CommandItem(CommandType.Workbook, WorkbookCommands.DeleteSheet,"Not needed for this command","Not needed for this command","Not needed for this command", "Name of worksheet being deleted","Not needed for this command"),
                new CommandItem(CommandType.Workbook, WorkbookCommands.MoveSheet,"See workbook options","Required if option selected is Before or After","Required if option selected is Before or After. Name or index of worksheet before or after position is relative to i.e. Sheet1 or 2.", "Name of worksheet to be moved","Not needed for this command"),
                new CommandItem(CommandType.Workbook, WorkbookCommands.CopySheet,"See workbook options","Not needed for this command","New name for worksheet being copied", "Name of worksheet being copied","Not needed for this command"),
                new CommandItem(CommandType.Workbook, WorkbookCommands.TargetSheet,"Not needed for this command","See reference options","Not needed for this command", "Worksheet being set to active. If 'By Name' is selected, enter the worksheet name. If 'By Index' is selected, enter the desired worksheet index to be targeted.","Not needed for this command"),
                new CommandItem(CommandType.Workbook, WorkbookCommands.RenameSheet,"Not needed for this command","See reference options","New name for worksheet", "Worksheet being renamed. If 'By Name' is selected, enter the worksheet name. If 'By Index' is selected, enter the desired worksheet index to be targeted.","Not needed for this command")
            };
        }

        public static List<CommandItem> GetRangeCommands()
        {
            return new List<CommandItem>
            {
                new CommandItem(CommandType.Range, RangeCommands.AddNamedRange,"See range options. Workbook scope will add a 'global' variable to the workbook while worksheet scope will add a 'local' variable to the current active worksheet. Workbook and worksheets do not currently support ranges with the same name. If workbook scope, avoid referencing multiple sheets under the same name.","Not needed for this command","Name of range being created", 
                "Desired cell range. Uses Excel cell reference style. For instance, 'B2' will select the cell one row down and one column over on the active worksheet and relative to the current selected cell. 'Commands!$H$10' will select H10 on the Commands worksheet. Easiest method I have found for entering valid ranges is pressing '=' in a cell and selecting a cell range. I then cut and paste everything excluding the '=' into the value cell. ","Not needed for this command"),
                new CommandItem(CommandType.Range, RangeCommands.RemoveNamedRange,"See range options. Workbook scope will attempt to remove the named range from the workbook scope while worksheet scope will attempt to remove the named range from the current worksheet's scope. If blank, will default to workbook scope. Workbook scope will currently look through all named ranges in workbook and worksheet and remove the first match it encounters.",
                "Not needed for this command","Not needed for this command", "Name of range being removed from name manager.","Not needed for this command"),
                new CommandItem(CommandType.Range, RangeCommands.SetNamedRange,"Not needed for this command","Not needed for this command", "Name of worksheet to be moved","See 'AddNamedRange' value column instructions.","Not needed for this command"),
                new CommandItem(CommandType.Range, RangeCommands.RenameRange,"See range options. Workbook scope will attempt to rename the named range from the workbook scope while worksheet scope will attempt to rename the named range from the current worksheet's scope. If blank, will default to workbook scope. Workbook scope will currently look through all named ranges in workbook and worksheet and rename the first match it encounters.",
                "Not needed for this command","New range name", "Old range name","Not needed for this command"),
                new CommandItem(CommandType.Range, RangeCommands.DeleteRangeContents,"Workbook scope will delete range contents within the workbook, whether it exist on the workbook or worksheet scope. Worksheet scope will delete range contents of a range that exists on the targeted sheet and within the worksheet scope. If left blank, will default to workbook scope but it is suggested to select worksheet scope for workbooks with a large set of named ranges for optimized performance.",
                "By name declares the name in the value column is a named range having its contents deleted. By index declares the value column is a cell array having its contents deleted","Not needed for this command", "Name of range having cells deleted.","Not needed for this command"),
                new CommandItem(CommandType.Range, RangeCommands.GetColumnRange,"Not needed for this command","Not needed for this command", "Not needed for this command","See 'AddNamedRange' value column instructions. Singular cell in desired column.","Returns the range value starting with the entered cell and ending with the cell before encountering a blank cell in the column."),
                new CommandItem(CommandType.Range, RangeCommands.GetRowRange,"Not needed for this command","Not needed for this command", "Not needed for this command","See 'AddNamedRange' value column instructions. Singular cell in the desired row","Returns the range range value starting with the entered cell and ending with the cell before encountering a blank cell in the row."),
                new CommandItem(CommandType.Range, RangeCommands.GetDataSetRange,"Not needed for this command","Not needed for this command", "Not needed for this command","See 'AddNamedRange' value column instructions. Singular cell located on the top leftmost location of a dataset.","Returns the range value starting with the entered cell and ending with the cell before encountering a blank cell in the last row and column."),
            };
        }      

        public static List<CommandItem> GetCodeCommands()
        {
            return new List<CommandItem>
            {
                new CommandItem(CommandType.Code, CodeCommands.Comment,"Not needed for this command. You are welcome to merge and center the options through value column on this row for cleaner comment readability.","Not needed for this command","Not needed for this command","Not needed for this command","Not needed for this command"),
                new CommandItem(CommandType.Code, CodeCommands.Sub,"Not needed for this command","Not needed for this command","Not needed for this command",
                "Name of subroutine being ran. This can be found on the named range located on the 'Command Type' header where a template is created. If subroutine is located on another worksheet, make sure to start value with 'SheetName' and '!' before the subroutine name. For example 'Sheet1!TheRoutineType'.","Not needed for this command"),
                new CommandItem(CommandType.Code, CodeCommands.If, "Not needed for this command","Not needed for this command","Not needed for this command","Logic using boolean test. Make sure value outputs 'TRUE' or 'YES' for true values. All other values will be treated as false.","Not needed for this command"),
                new CommandItem(CommandType.Code, CodeCommands.EndIf, "Not needed for this command","Not needed for this command","Not needed for this command","Not needed for this command","Not needed for this command"),
                new CommandItem(CommandType.Code, CodeCommands.Repeat,"Not needed for this command. You are welcome to merge and center the options through value column on this row for cleaner comment readability.","Not needed for this command","Not needed for this command","Number of repetitions desired (must be an integer value)","Outputs the counter value for the repeat. First value outputted will be 1 and will increment by 1 until the number in the value column is reached."),
                new CommandItem(CommandType.Code, CodeCommands.EndRepeat,"Not needed for this command.","Not needed for this command","Not needed for this command","Not needed for this command","Not needed for this command"),
                new CommandItem(CommandType.Code, CodeCommands.Stop,"Not needed for this command.","Not needed for this command","Not needed for this command","Not needed for this command","Not needed for this command"),
            };
        }

        public static List<CommandItem> GetWorksheetCommands()
        {
            return new List<CommandItem>
            {
                new CommandItem(CommandType.Worksheet, WorksheetCommands.AddColumn,"Not needed for this command","Not needed for this command","Not needed for this command","Cell to have column added to. Targeted cell column will shift to the right.","Not needed for this command"),
                new CommandItem(CommandType.Worksheet, WorksheetCommands.AddRow,"Not needed for this command","Not needed for this command","Not needed for this command","Cell to have row added to. Targeted cell row will shift down.","Not needed for this command"),
                new CommandItem(CommandType.Worksheet, WorksheetCommands.MoveColumn,"Not needed for this command","Not needed for this command","Cell to have column moved left of. For example, entering cell 'F1' will insert the new column in column E.", "Cell to be moved. Enter a single cell in the desired column", "Not needed for this command"),
                new CommandItem(CommandType.Worksheet, WorksheetCommands.MoveRow,"Not needed for this command","Not needed for this command","Cell to have row moved above. For example, entering cell 'F3' will insert the new row in row 2.", "Cell to be moved. Enter a single cell in the desired row", "Not needed for this command"),
                new CommandItem(CommandType.Worksheet, WorksheetCommands.DeleteColumn,"Not needed for this command","Not needed for this command","Not needed for this command","Cell to have column deleted. Columns to the right of the column being deleted will shift left.","Not needed for this command"),
                new CommandItem(CommandType.Worksheet, WorksheetCommands.DeleteRow,"Not needed for this command","Not needed for this command","Not needed for this command","Cell to have row deleted. Rows to below the row being deleted will shift up.","Not needed for this command"),
            };
        }

        public static List<CommandItem> GetFilterCommands()
        {
            return new List<CommandItem>
            {
                 new CommandItem(CommandType.Filter, FilterCommands.AddDataFilter,"Specifies type of filter added." + "\n" + "AND: Logical 'AND' of criteria 1 and 2." + "\n" + "OR: Logical 'OR' of criteria 1 and 2." + "\n" + "FILTER VALUES: filters data for values entered. Read on in named column for advanced filtering." + "\n" +
                 "TOP 10 ITEMS and BOTTOM 10 ITEMS: Gets the top or bottom 10 items of the data specified either sorted numerically or alphabetically depending on data in the column. You can enter other integer values in the name column to choose the top 5 items, for instance." + "\n" +
                 "TOP 10 PERCENT and BOTTOM 10 PERCENT: Gets the top or bottom 10 percent of the data specified either sorted numerically or alphabetically depending on data in the column. You can enter other integer values in the name column to choose the top 5 percent, for instance.",
                 "Column index, starting from the left and at 1, you wish to apply the filter to",
                 "Search criteria for the current filter." + "\n" +  "If you have multiple values to filter for, use comma delimited text with no spaces to break up values i.e. 'Car,Truck,Van'." + "\n" +  "If you are looking for numerics or dates, you can use the '<,<=,>,>=, and <>' operators to find values related to the desired value." + "\n" + 
                 "If you are looking for blank cells, just place an '=' in the name column." + "\n" + "If you are looking for nonblank fields, enter '<>' in the name column." + "\n" + "You can also use wildcards in your filter criteria. You can search using a '?' to fill in a single character." +
                 " For example, 'sm?th' finds 'smith' and 'smyth'." + "\n" + "You can also use an asterisk to represent any number of characters. For example, '*east' finds 'Northeast' and 'Southeast'.","Cell range to apply filtering to. Include the header row in the range.","Second set of criteria, if any, to apply for filter. If 'AND' or 'OR' options are selected, this is where you will apply your comparative criteria."),
                 new CommandItem(CommandType.Filter, FilterCommands.ClearDataFilters,"Not needed for this command","Not needed for this command","Not needed for this command","Not needed for this command","Not needed for this command"),
                 new CommandItem(CommandType.Filter, FilterCommands.DeleteDataFilters,"Not needed for this command","Not needed for this command","Not needed for this command","Not needed for this command","Not needed for this command"),
            };
        }

        public static List<CommandItem> GetDataCommands()
        {
            return new List<CommandItem>
            {
                new CommandItem(CommandType.Data, DataCommands.SetValue,"Not needed for this command", "Not needed for this command", "Not needed for this command", "Not needed for this command", "Not needed for this command"),                
                new CommandItem(CommandType.Data, DataCommands.FindAndReplace,"See match value options. Sets the search parameters to the whole cell value or a part of it.", "Set true to match case of search. False will not match the case.", "Text to search for", "Cell range find and replace is to be conducted in.", "Value to replace finding(s)."),
            };
        }
    }
}
