using Exceleration.Helpers.Enums;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Exceleration.Helpers
{
    public static class ReflectionHelper
    {
        public static Type GetTypeFromEnum(TypeEnum typeEnum)
        {
            switch(typeEnum)
            {
                case TypeEnum.String:
                    return Type.GetType("System.String");
                case TypeEnum.Integer:
                    return Type.GetType("System.Int32");
                case TypeEnum.Decimal:
                    return Type.GetType("System.Decimal");
                default:
                    throw new Exception("Could not return a type that matches the parameter");
            }
        }
    }
}
