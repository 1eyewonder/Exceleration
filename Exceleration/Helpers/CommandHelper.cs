using Exceleration.Options;
using Exceleration.DTOS;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Exceleration
{
    public static class CommandHelper
    {
        public static List<CommandItem> GetSheetCommands()
        {
            return new List<CommandItem>
            {
                new CommandItem(SheetCommands.AddSheet,"See sheet options","Option for reference target","Target Name, if any",null),
                new CommandItem(SheetCommands.DeleteSheet,"See sheet options","Option for reference target","Target Name, if any",null),
                new CommandItem(SheetCommands.MoveSheet,"See sheet options","Option for reference target","Target Name, if any",null),
                new CommandItem(SheetCommands.CopySheet,"See sheet options","Option for reference target","Target Name, if any",null),
            };
        }
    }
}
