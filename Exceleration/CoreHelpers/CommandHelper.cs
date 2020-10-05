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
                new CommandItem(CommandType.Workbook, WorkbookCommands.CopySheet,"See workbook options","Not required for this command","New name for worksheet being copied", "Name of worksheet being copied"),
                new CommandItem(CommandType.Workbook, WorkbookCommands.TargetSheet,"See workbook options","Not required for this command","New name for worksheet", "Worksheet being renamed"),
                new CommandItem(CommandType.Workbook, WorkbookCommands.RenameSheet,"See workbook options","Not required for this command","New name for worksheet being copied", "Name of worksheet being copied")
            };
        }

        public static List<CommandItem> GetRangeCommands()
        {
            return new List<CommandItem>
            {
                new CommandItem(CommandType.Range, RangeCommands.AddNamedRange,"Not needed for this command","Not needed for this command","Not needed for this command", "Name of new range being added"),
                new CommandItem(CommandType.Range, RangeCommands.RemoveNamedRange,"Not needed for this command","Not needed for this command","Not needed for this command", "Name of range being removed from name manager"),
                new CommandItem(CommandType.Range, RangeCommands.SetNamedRange,"Not needed for this command","Required if option selected is Before or After","Required if option selected is Before or After. Name or index of worksheet before or after position is relative to i.e. Sheet1 or 2.", "Name of worksheet to be moved"),
                new CommandItem(CommandType.Range, RangeCommands.RenameRange,"Not needed for this command","Not required for this command","New range name", "Old range name"),
                new CommandItem(CommandType.Range, RangeCommands.DeleteRangeContents,"Not needed for this command","Not required for this command","Not needed for this command", "Name of range having cells deleted")
            };
        }
    }
}
