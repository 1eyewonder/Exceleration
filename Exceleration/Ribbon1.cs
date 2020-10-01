using System;
using System.Linq;
using System.Windows.Forms;
using Exceleration.Options;
using Exceleration.Helpers;
using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;
using Exceleration.Forms;
using Exceleration.Commands;
using Exceleration.Build;
using Exceleration.CoreHelpers;

namespace Exceleration
{
    public partial class Ribbon1
    {
        private MainBuilder _mainBuilder;
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

        /// <summary>
        /// Adds command worksheet to workbook. Used for general command explanations and option lists
        /// </summary>
        private void AddCommands()
        {
            // Attempts to find a currently existing Command worksheet
            Excel.Workbook workbook = Globals.ThisAddIn.Application.ActiveWorkbook;
            var workSheet = workbook.GetWorksheets().FirstOrDefault(x => x.Name == "Commands");

            // Adds a worksheet named Commands if it does not already exist
            workSheet = workbook.CreateNewWorksheet("Commands");

            // Adds column headers to command table
            workSheet.Range["A1"].Value = "Command Type";
            workSheet.Range["B1"].Value = "Command";
            workSheet.Range["C1"].Value = "Options";
            workSheet.Range["D1"].Value = "Reference";
            workSheet.Range["E1"].Value = "Name";
            workSheet.Range["F1"].Value = "Value";

            #region SheetOptions
            // Creates sheet options table
            workSheet.Range["J1"].Value = "Sheet Options";
            var sheetOptions = OptionHelper.GetSheetOptions();

            int j = 2;
            foreach (var o in sheetOptions)
            {
                workSheet.Range[$"J{j}"].Value = o;
                j++;
            }

            // Gets count of fields to assign range for naming
            var sheetOptionsCount = sheetOptions.Count;
            workSheet.CreateNamedRange(nameof(SheetOptions), $"J2:J{1 + sheetOptionsCount}");
            #endregion

            #region ReferenceOptions
            // Creates sheet options table
            workSheet.Range["K1"].Value = "Reference Options";
            var referenceOptions = OptionHelper.GetReferenceOptions();

            int q = 2;
            foreach (var o in referenceOptions)
            {
                workSheet.Range[$"K{q}"].Value = o;
                q++;
            }

            // Gets count of fields to assign range for naming
            var referenceOptionsCount = referenceOptions.Count;
            workSheet.CreateNamedRange(nameof(ReferenceOptions), $"K2:K{1 + referenceOptionsCount}");
            #endregion

            #region Commands
            // Gets commands for each category of methods
            var sheetCommands = CommandHelper.GetSheetCommands();

            // Starts counter at 2 due to Excel row we are starting at
            var i = 2;

            foreach (var c in sheetCommands)
            {
                workSheet.Range[$"A{i}"].Value = c.CommandType;
                workSheet.Range[$"B{i}"].Value = c.Command;
                workSheet.Range[$"C{i}"].Value = c.Options;
                workSheet.Range[$"D{i}"].Value = c.Reference;
                workSheet.Range[$"E{i}"].Value = c.Name;
                workSheet.Range[$"F{i}"].Value = c.Value;
                i++;
            }

            // Gets count of fields to assign range for naming
            var commandCount = sheetCommands.Count;
            workSheet.CreateNamedRange("Commands", $"B2:B{1 + commandCount}");
            #endregion

            #region Styling
            // Styles the command table
            var stylingRange = (Excel.Range)workSheet.Range["A:F"];
            stylingRange.ColumnWidth = 25;
            stylingRange.Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            stylingRange.Cells.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            stylingRange.Cells.WrapText = true;

            // Selects and styles the column headers
            var topRange = workSheet.Range["A1:F1"];
            topRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.CornflowerBlue);
            topRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
            topRange.Font.Bold = true;

            // Styles the options tables
            stylingRange = (Excel.Range)workSheet.Range["J:K"];
            stylingRange.ColumnWidth = 25;
            stylingRange.Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;    
            #endregion
        }

        private void AddTemplate()
        {
            Excel.Worksheet workSheet = null;
            Excel.Application XlApp = Globals.ThisAddIn.Application;

            // Checks to see if the current worksheet is empty
            var empty = ((Excel.Worksheet)XlApp.ActiveSheet).WorkSheetEmpty();

            // If empty
            if (empty)
            {
                workSheet = XlApp.ActiveSheet;
            }

            // If not empty, create new worksheet
            else
            {
                InputForm form = new InputForm("Enter new worksheet name");

                form.ShowDialog();

                var workSheetName = form.TextInput;

                if (string.IsNullOrEmpty(workSheetName))
                {
                    return;
                }

                XlApp.ActiveWorkbook.CreateNewWorksheet(workSheetName);
                XlApp.ActiveWorkbook.ActivateSheet(workSheetName);
                workSheet = XlApp.ActiveSheet;
            }

            // Adds column headers to template table
            workSheet.Range["A5"].Value = "Command Type";
            workSheet.Range["B5"].Value = "Command";
            workSheet.Range["C5"].Value = "Options";
            workSheet.Range["D5"].Value = "Reference";
            workSheet.Range["E5"].Value = "Name";
            workSheet.Range["F5"].Value = "Value";

            workSheet.Range["A5"].Name = $"{workSheet.Name}Type";

            // Selects and styles the column headers
            var topRange = workSheet.Range["A5:F5"];
            topRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.CornflowerBlue);
            topRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
            topRange.Font.Bold = true;

            // Styles the command columns
            var stylingRange = (Excel.Range)workSheet.Range["A:F"];
            stylingRange.ColumnWidth = 25;
            stylingRange.Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
        }

        private void AddSheetCommands(object sender, RibbonControlEventArgs e)
        {                    
            Excel.Workbook workbook = Globals.ThisAddIn.Application.ActiveWorkbook;
            Excel.Worksheet worksheet = workbook.ActiveSheet;
            Excel.Range range = Globals.ThisAddIn.Application.ActiveCell;

            //Checks if commands page exists
            if (!workbook.WorkSheetExists("Commands"))
            {
                MessageBox.Show("Please run 'Add Commands' and then try again.");
            }

            // Get row and next column indices
            var nextColumn = WorksheetHelper.GetColumnName(range.Column+1);
            var thisRow = range.Row;

            // Command type column
            range.Value = CommandType.Worksheet;

            // Command column
            range = worksheet.Range[$"{nextColumn}" + $"{thisRow}"];
            range.AddDropDownList("Commands");

            // Gets next column index
            nextColumn = WorksheetHelper.GetColumnName(range.Column + 1);

            // Options column
            range = worksheet.Range[$"{nextColumn}" + $"{thisRow}"];
            range.AddDropDownList(nameof(SheetOptions));

            // Gets next column index
            nextColumn = WorksheetHelper.GetColumnName(range.Column + 1);

            // Reference column
            range = worksheet.Range[$"{nextColumn}" + $"{thisRow}"];
            range.AddDropDownList(nameof(ReferenceOptions));
        }

        private void AddTemplateButton_Click(object sender, RibbonControlEventArgs e)
        {
            AddTemplate();
        }

        private void RunButton_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Worksheet workSheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
            _mainBuilder = new MainBuilder(workSheet);

            try
            {
                var logs = _mainBuilder.Run(true);

                if (logs.Count > 0)
                {
                    MessageBox.Show("Done", "Application Complete - Please Review Logs", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }
                else
                {
                    MessageBox.Show("Done", "Application Complete", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ran into an error while running", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
