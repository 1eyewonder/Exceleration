using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using Exceleration.Helpers.Extensions;
using Exceleration.Build;
using Exceleration.Helpers;

namespace Exceleration.CoreHelpers
{
    public class ObjectMapHelper : ExcelParse
    {
        public ObjectMapHelper(Excel.Worksheet worksheet) : base(worksheet)
        {

        }

        public Dictionary<string, Type> GetObjectMap(string rangeName)
        {
            Excel.Range startCell = null;
            Excel.Workbook workbook = Globals.ThisAddIn.Application.ActiveWorkbook;
            var returnDictionary = new Dictionary<string, Type>();

            try
            {
                // Looks for range to start on
                if (!string.IsNullOrEmpty(rangeName))
                {
                    startCell = workbook.GetWorksheet("Object Maps").Range[rangeName];
                }

                if (startCell == null) throw new Exception("Could not find start range");
            }
            catch
            {
                throw new ArgumentException("Issue finding starting code cell. Please check that the workbook has an object map worksheet.");
            }

            var nameColumn = startCell.Column;
            var typeColumn = nameColumn + 1;

            string name;
            Type type;

            int i = startCell.Row;

            while (!string.IsNullOrEmpty(_getString(i, nameColumn)))
            {
                name = _getString(i, nameColumn);

                var typeEnum = EnumHelper.GetEnumFromString(_getString(i, typeColumn));
                type = ReflectionHelper.GetTypeFromEnum(typeEnum);

                returnDictionary.Add(name, type);

                i++;
            }

            return returnDictionary;
        }
    }
}
