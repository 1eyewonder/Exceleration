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
        public static List<CommandItem> GetSheetCommands()
        {
            return new List<CommandItem>
            {
                new CommandItem(CommandType.Worksheet, SheetCommands.AddSheet,"Not needed for this command","Not needed for this command","Not needed for this command", "Name of new worksheet being added"),
               new CommandItem(CommandType.Worksheet, SheetCommands.DeleteSheet,"Not needed for this command","Not needed for this command","Not needed for this command", "Name of new worksheet being added"),
                new CommandItem(CommandType.Worksheet, SheetCommands.MoveSheet,"See sheet options","Required if option selected is Before or After","Required if option selected is Before or After. Name or index of worksheet before or after position is relative to i.e. Sheet1 or 2.", "Name of worksheet to be moved"),
                new CommandItem(CommandType.Worksheet, SheetCommands.CopySheet,"See sheet options","Not required for this command","New name for worksheet being copied", "Name of worksheet being copied")
            };
        }
    }
}
