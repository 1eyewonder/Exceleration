using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Exceleration.DTOS
{
    public class CommandItem
    {
        public string Command { get; set; }
        public string Options { get; set; }
        public string Reference { get; set; }
        public string Name { get; set; }
        public string Value { get; set; }

        public CommandItem(string command, string options, string reference, string name, string value)
        {
            Command = command;
            Options = options;
            Reference = reference;
            Name = name;
            Value = value;
        }
    }
}
