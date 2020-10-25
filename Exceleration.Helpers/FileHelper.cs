using System;
using System.Data;
using System.IO;
using System.Text.RegularExpressions;

namespace Exceleration.Helpers
{
    public static class FileHelper
    {
        /// <summary>
        /// Checks if a string contains a valid file path
        /// </summary>
        /// <param name="path">Target file path</param>
        /// <returns></returns>
        public static bool IsValidPath(string path)
        {
            Regex driveCheck = new Regex(@"^[a-zA-Z]:\\$");
            if (!driveCheck.IsMatch(path.Substring(0, 3))) return false;
            string strTheseAreInvalidFileNameChars = new string(Path.GetInvalidPathChars());
            strTheseAreInvalidFileNameChars += @":/?*" + "\"";
            Regex containsABadCharacter = new Regex("[" + Regex.Escape(strTheseAreInvalidFileNameChars) + "]");

            if (containsABadCharacter.IsMatch(path.Substring(3, path.Length - 3)))
            {
                return false;
            }

            return true;
        }

        /// <summary>
        /// Copies template to target path and removes read only properties, if applicable
        /// </summary>
        /// <param name="oldPath"></param>
        /// <param name="newPath"></param>
        /// I use this to copy read-only excel templates to create editable versions
        public static void CopyTemplate(string oldPath, string newPath)
        {
            //Removes the readonly from the old file to prevent errors
            FileAttributes attributes = File.GetAttributes(oldPath);
            if ((attributes & FileAttributes.ReadOnly) == FileAttributes.ReadOnly)
            {
                // Make the file RW
                attributes = RemoveAttribute(attributes, FileAttributes.ReadOnly);
                File.SetAttributes(oldPath, attributes);
                Console.WriteLine("The {0} file is no longer RO.", newPath);
            }

            //Opens the template file and allows overwrite
            File.Copy(oldPath, newPath, true);
        }

        /// <summary>
        /// Returns a list of file attributes to remove
        /// </summary>
        /// <param name="attributes"></param>
        /// <param name="attributesToRemove"></param>
        /// <returns></returns>
        private static FileAttributes RemoveAttribute(FileAttributes attributes, FileAttributes attributesToRemove)
        {
            return attributes & ~attributesToRemove;
        }     

        /// <summary>
        /// Reads data from a csv file and converts it to
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="delimiter"></param>
        /// <returns></returns>
        public static DataTable GetDataTableFromFile(string filePath, string delimiter, string fileType = ".csv")
        {
            DataTable dataTable = new DataTable();

            if (IsValidPath(filePath) && filePath.Substring(filePath.Length - fileType.Length) == fileType)
            {
                using (StreamReader reader = new StreamReader(filePath))
                {
                    // Used none enum in case csv data contains a blank value
                    string[] headers = reader.ReadLine().Split(new string[] { delimiter }, StringSplitOptions.None);

                    foreach (string header in headers)
                    {
                        dataTable.Columns.Add(header);
                    }

                    while (!reader.EndOfStream)
                    {
                        string[] rows = reader.ReadLine().Split(new string[] { delimiter }, StringSplitOptions.None);
                        DataRow row = dataTable.NewRow();

                        for (int i = 0; i < headers.Length; i++)
                        {
                            row[i] = rows[i];
                        }

                        dataTable.Rows.Add(row);
                    }

                    return dataTable;
                }                
            }
            else
            {
                throw new Exception("File path was not valid. Please try again");
            }
        }
    }
}
