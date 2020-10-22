using Exceleration.Helpers.Enums;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Exceleration.Helpers
{
    public static class EnumHelper
    {
        public static TypeEnum GetEnumFromString(string text)
        {
            switch(text.ToUpper())
            {
                case "STRING":
                    return TypeEnum.String;
                case "INTEGER":
                    return TypeEnum.Integer;
                case "DECIMAL":
                    return TypeEnum.Decimal;
                default:
                    throw new Exception();
            }
        }
    }
}
