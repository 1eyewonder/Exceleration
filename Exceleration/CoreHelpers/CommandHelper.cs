using Exceleration.Options;
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
                new CommandItem(CommandType.Workbook, WorkbookCommands.AddSheet,"Not needed for this command","Not needed for this command","Not needed for this command", "Name of new worksheet being added"),
                new CommandItem(CommandType.Workbook, WorkbookCommands.DeleteSheet,"Not needed for this command","Not needed for this command","Not needed for this command", "Name of worksheet being deleted"),
                new CommandItem(CommandType.Workbook, WorkbookCommands.MoveSheet,"See workbook options","Required if option selected is Before or After","Required if option selected is Before or After. Name or index of worksheet before or after position is relative to i.e. Sheet1 or 2.", "Name of worksheet to be moved"),
                new CommandItem(CommandType.Workbook, WorkbookCommands.CopySheet,"See workbook options","Not needed for this command","New name for worksheet being copied", "Name of worksheet being copied"),
                new CommandItem(CommandType.Workbook, WorkbookCommands.TargetSheet,"Not needed for this command","See reference options","Not needed for this command", "Worksheet being set to active. If 'By Name' is selected, enter the worksheet name. If 'By Index' is selected, enter the desired worksheet index to be targeted."),
                new CommandItem(CommandType.Workbook, WorkbookCommands.RenameSheet,"Not needed for this command","See reference options","New name for worksheet", "Worksheet being renamed. If 'By Name' is selected, enter the worksheet name. If 'By Index' is selected, enter the desired worksheet index to be targeted.")
            };
        }

        public static List<CommandItem> GetRangeCommands()
        {
            return new List<CommandItem>
            {
                new CommandItem(CommandType.Range, RangeCommands.AddNamedRange,"See range options. Workbook scope will add a 'global' variable to the workbook while worksheet scope will add a 'local' variable to the current active worksheet. Workbook and worksheets do not currently support ranges with the same name.","Not needed for this command","Name of range being created", 
                "Desired cell range. Uses Excel cell reference style. For instance, 'B2' will select the cell one row down and one column over on the active worksheet and relative to the current selected cell. 'Commands!$H$10' will select H10 on the Commands worksheet. Easiest method I have found for entering valid ranges is pressing '=' in a cell and selecting a cell range. I then cut and paste everything excluding the '=' into the value cell. "),
                new CommandItem(CommandType.Range, RangeCommands.RemoveNamedRange,"See range options. Workbook scope will attempt to remove the named range from the workbook scope while worksheet scope will attempt to remove the named range from the current worksheet's scope. If blank, will default to workbook scope. Workbook scope will currently look through all named ranges in workbook and worksheet and remove the first match it encounters.",
                "Not needed for this command","Not needed for this command", "Name of range being removed from name manager."),
                new CommandItem(CommandType.Range, RangeCommands.SetNamedRange,"Not needed for this command","Not needed for this command", "Name of worksheet to be moved","See 'AddNamedRange' value column instructions."),
                new CommandItem(CommandType.Range, RangeCommands.RenameRange,"See range options. Workbook scope will attempt to rename the named range from the workbook scope while worksheet scope will attempt to rename the named range from the current worksheet's scope. If blank, will default to workbook scope. Workbook scope will currently look through all named ranges in workbook and worksheet and rename the first match it encounters.",
                "Not needed for this command","New range name", "Old range name"),
                new CommandItem(CommandType.Range, RangeCommands.DeleteRangeContents,"Not needed for this command","Not needed for this command","Not needed for this command", "Name of range having cells deleted.")
            };
        }
    }
}
