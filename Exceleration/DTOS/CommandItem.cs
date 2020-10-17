using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Exceleration.DTOS
{
    public class CommandItem
    {
        public string CommandType { get; set; }
        public string Command { get; set; }
        public string Options { get; set; }
        public string Reference { get; set; }
        public string Name { get; set; }
        public string TargetValue { get; set; }
        public string AuxillaryValue { get; set; }

        public CommandItem(string commandType, string command, string options, string reference, string name, string targetValue, string auxillaryValue)
        {
            CommandType = commandType;
            Command = command;
            Options = options;
            Reference = reference;
            Name = name;
            TargetValue = targetValue;
            AuxillaryValue = auxillaryValue;
        }
    }
}
