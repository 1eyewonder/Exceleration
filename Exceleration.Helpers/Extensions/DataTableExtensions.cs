using Exceleration.Helpers.Enums;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Exceleration.Helpers.Extensions
{
    public static class DataTableExtensions
    {
        /// <summary>
        /// Writes data from data table to csv file
        /// </summary>
        /// <param name="dataTable">Target data table</param>
        /// <param name="filePath">File path</param>
        /// <param name="delimiter">Text delimiter</param>
        /// I left the delimiter vague so I wouldn't have to maintain this method everytime someone had a special delimiter
        public static void WriteToCsvFile(this DataTable dataTable, string filePath, string delimiter)
        {
            StringBuilder fileContent = new StringBuilder();

            foreach (var col in dataTable.Columns)
            {
                fileContent.Append(col.ToString() + $"{delimiter}");
            }

            fileContent.Replace($"{delimiter}", Environment.NewLine, fileContent.Length - delimiter.Length, delimiter.Length);            

            foreach (DataRow dr in dataTable.Rows)
            {
                foreach (var column in dr.ItemArray)
                {
                    fileContent.Append("\"" + column.ToString() + $"\"{delimiter}");
                }

                fileContent.Replace($"{delimiter}", Environment.NewLine, fileContent.Length - delimiter.Length, delimiter.Length);                                   
            }

            if (FileHelper.IsValidPath(filePath))
            {
                System.IO.File.WriteAllText(filePath, fileContent.ToString());
            }
            else
            {
                throw new Exception("Please enter a valid path to save the csv file");
            }            
        }

        public static string WriteToJson(this DataTable dataTable)
        {
            return JsonConvert.SerializeObject(dataTable);
        }
    }
}
