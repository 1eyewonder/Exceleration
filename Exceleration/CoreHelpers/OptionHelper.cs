using Exceleration.Options;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Reflection;
using Exceleration.Helpers.Enums;
using Excel = Microsoft.Office.Interop.Excel;

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

        public static List<string> GetExcelAutoFilterOptions()
        {
            var options = new ExcelAutoFilterOptions();
            return options.GetFieldValues();
        }

        public static List<string> GetMatchValueOptions()
        {
            var options = new MatchValueOptions();
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

        public static Excel.XlAutoFilterOperator GetExcelAutoFilterOperatorFromString(string text)
        {
            switch (text)
            {
                case ExcelAutoFilterOptions.And:
                    return Excel.XlAutoFilterOperator.xlAnd;

                case ExcelAutoFilterOptions.Or:
                    return Excel.XlAutoFilterOperator.xlOr;

                case ExcelAutoFilterOptions.Top10Items:
                    return Excel.XlAutoFilterOperator.xlTop10Items;

                case ExcelAutoFilterOptions.Top10Percent:
                    return Excel.XlAutoFilterOperator.xlTop10Percent;

                case ExcelAutoFilterOptions.Bottom10Items:
                    return Excel.XlAutoFilterOperator.xlBottom10Items;

                case ExcelAutoFilterOptions.Bottom10Percent:
                    return Excel.XlAutoFilterOperator.xlBottom10Percent;

                case ExcelAutoFilterOptions.FilterValues:
                    return Excel.XlAutoFilterOperator.xlFilterValues;

                default:
                    return Excel.XlAutoFilterOperator.xlAnd;
            }
        }
    }
}
