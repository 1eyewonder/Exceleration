using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using Exceleration.Helpers.Extensions;
using Exceleration.CoreHelpers;
using Exceleration.Helpers;

namespace Exceleration.Commands
{
    public class CsvCommands
    {
        public const string WriteToCsv = "WRITE TO CSV";
        public const string ReadFromCsv = "READ FROM CSV";

        public void WriteToCsvCommand(Excel.Workbook workbook, Excel.Worksheet activeSheet, string objectMap, string dataRange, string csvFilePath, string delimiter)
        {
            Excel.Worksheet worksheet = workbook.GetWorksheet("Object Maps");
            var objectMapHelper = new ObjectMapHelper(worksheet);

            if (string.IsNullOrEmpty(delimiter))
            {
                delimiter = ",";
            }

            if (worksheet.NamedRangeExists(objectMap))
            {
                var dictionary = objectMapHelper.GetObjectMap(objectMap);

                if (activeSheet.IsRange(dataRange))
                {
                    var someTable = activeSheet.Range[dataRange].ConvertToDataTable(dictionary);
                    someTable.WriteToCsvFile(csvFilePath, delimiter);
                }
            }
        }

        public void ReadFromCsvCommand(Excel.Worksheet worksheet, string filePath, string delimiter, string range)
        {
            if (string.IsNullOrEmpty(delimiter))
            {
                delimiter = ",";
            }

            var table = FileHelper.GetDataTableFromFile(filePath, delimiter);

            if (worksheet.IsRange(range) || worksheet.NamedRangeExists(range))
            {
                worksheet.WriteFromDataTable(table, worksheet.Range[range]);
            }            
         
            Console.WriteLine("Hello");
        }
    }
}
