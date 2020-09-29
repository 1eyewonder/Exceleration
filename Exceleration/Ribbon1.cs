using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Exceleration.Commands;
using Exceleration.Helpers;
using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;

namespace Exceleration
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Workbook workbook = Globals.ThisAddIn.Application.ActiveWorkbook;

            //MessageBox.Show("Testing","Testing The Button", MessageBoxButtons.OK);
            AddCommands();
            workbook.WorkSheetExists("testing");
        }

        private void AddCommands()
        {
            Excel.Workbook workbook = Globals.ThisAddIn.Application.ActiveWorkbook;
            var workSheet = workbook.GetWorksheets().FirstOrDefault(x => x.Name == "Commands");

            if (workSheet == null)
            {
                workbook.Worksheets.Add();

                workSheet = Globals.ThisAddIn.Application.ActiveSheet;

                workSheet.Name = "Commands";
            }

            workSheet.Range["A1"].Value = "Command";
            workSheet.Range["B1"].Value = "Options";
            workSheet.Range["C1"].Value = "Reference";
            workSheet.Range["D1"].Value = "Name";
            workSheet.Range["E1"].Value = "Value";
            workSheet.Range["F1"].Value = "Notes";
        }
    }
}
