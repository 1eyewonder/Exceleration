using Exceleration.Options;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Reflection;

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

        public static List<string> GetSheetOptions()
        {
            var options = new SheetOptions();
            return options.GetFieldValues();
        }   

        public static List<string> GetReferenceOptions()
        {
            var options = new ReferenceOptions();
            return options.GetFieldValues();
        }
    }
}
