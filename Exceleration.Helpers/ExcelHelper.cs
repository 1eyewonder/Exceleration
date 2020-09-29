using System;
using System.Diagnostics;
using Excel = Microsoft.Office.Interop.Excel;

namespace Exceleration.Helpers
{
    public static class ExcelHelper
    {
        /// <summary>
        /// Method used to close out Excel application instances to prevent "ghost" instances found in task manager background
        /// </summary>
        /// <param name="obj"></param>
        private static void ReleaseObject(this Object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
            }
            catch (Exception ex)
            {
                obj = null;
                Debug.Print(ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }

        /// <summary>
        /// Quits Excel and clears instance from memory
        /// </summary>
        /// <param name="app"></param>
        public static void QuitExcel(this Excel.Application app)
        {
            app.Quit();
            ReleaseObject(app);
        }
    }
}
