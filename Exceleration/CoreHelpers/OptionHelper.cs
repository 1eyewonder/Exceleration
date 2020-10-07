using Exceleration.Options;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Reflection;
using Exceleration.Helpers.Enums;

namespace Exceleration.CoreHelpers
{
    public static class OptionHelper
    {
        /// <summary>
        /// Returns all field values as a list of string values
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="obj"></param>
        /// <returns></returns>
        private static List<string> GetFieldValues<T>(this T obj)
        {
            var optionsList = new List<string>();

            foreach (FieldInfo field in obj.GetType().GetFields())
            {
                optionsList.Add(field.GetValue(obj).ToString());
            }

            return optionsList;
        }

        public static List<string> GetWorkbookOptions()
        {
            var options = new WorkbookOptions();
            return options.GetFieldValues();
        }   

        public static List<string> GetReferenceOptions()
        {
            var options = new ReferenceOptions();
            return options.GetFieldValues();
        }

        public static List<string> GetRangeOptions()
        {
            var options = new RangeOptions();
            return options.GetFieldValues();
        }

        public static ReferenceEnum GetReferenceEnumFromString(string reference)
        {
            switch (reference)
            {
                case (ReferenceOptions.ByName):
                    return ReferenceEnum.ByName;

                case (ReferenceOptions.ByIndex):
                    return ReferenceEnum.ByIndex;

                default:
                    return ReferenceEnum.ByName;
            }
        }

        public static PositionalEnum GetPositionalEnumFromString(string position)
        {
            switch (position)
            {
                case (WorkbookOptions.After):
                    return PositionalEnum.After;

                case (WorkbookOptions.AtBeginning):
                    return PositionalEnum.AtBeginning;

                case (WorkbookOptions.AtEnd):
                    return PositionalEnum.AtEnd;

                case (WorkbookOptions.Before):
                    return PositionalEnum.Before;

                default:
                    return PositionalEnum.AtEnd;
            }
        }
    }
}
