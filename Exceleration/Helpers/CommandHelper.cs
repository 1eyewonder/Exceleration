using Exceleration.Commands;
using Exceleration.DTOS;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Exceleration
{
    public class CommandHelper
    {
        public List<CommandItem> GetSheetCommands()
        {
            return new List<CommandItem>
            {
                new CommandItem(SheetCommand.AddSheet, "Adds a new worksheet","Option for reference target","Target Name, if any",null)
            };
        }
    }
}
