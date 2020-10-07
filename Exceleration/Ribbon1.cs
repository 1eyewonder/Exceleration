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
using Exceleration.Helpers.Extensions;

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
            AddCommands();
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
            workSheet.Range["J1"].Value = nameof(WorkbookOptions);
            var sheetOptions = OptionHelper.GetWorkbookOptions();

            int j = 2;
            foreach (var o in sheetOptions)
            {
                workSheet.Range[$"J{j}"].Value = o;
                j++;
            }

            // Gets count of fields to assign range for naming
            var sheetOptionsCount = sheetOptions.Count;
            workbook.CreateNamedRange(nameof(WorkbookOptions), $"J2:J{1 + sheetOptionsCount}");
            #endregion

            #region ReferenceOptions
            // Creates sheet options table
            workSheet.Range["K1"].Value = nameof(ReferenceOptions);
            var referenceOptions = OptionHelper.GetReferenceOptions();

            int q = 2;
            foreach (var o in referenceOptions)
            {
                workSheet.Range[$"K{q}"].Value = o;
                q++;
            }

            // Gets count of fields to assign range for naming
            var referenceOptionsCount = referenceOptions.Count;
            workbook.CreateNamedRange(nameof(ReferenceOptions), $"K2:K{1 + referenceOptionsCount}");
            #endregion

            #region RangeOptions
            // Creates range options table
            workSheet.Range["L1"].Value = nameof(RangeOptions);
            var rangeOptions = OptionHelper.GetRangeOptions();

            int p = 2;
            foreach (var o in rangeOptions)
            {
                workSheet.Range[$"L{p}"].Value = o;
                p++;
            }

            // Gets count of fields to assign range for naming
            var rangeOptionsCount = rangeOptions.Count;
            workbook.CreateNamedRange(nameof(RangeOptions), $"L2:L{1 + rangeOptionsCount}");
            #endregion

            #region Workbook Commands
            // Gets commands for each category of methods
            var workbookCommands = CommandHelper.GetWorkbookCommands();

            // Starts counter at 2 due to Excel row we are starting at
            var i = 2;

            foreach (var c in workbookCommands)
            {
                workSheet.Range[$"A{i}"].Value = c.CommandType;
                workSheet.Range[$"B{i}"].Value = c.Command;
                workSheet.Range[$"C{i}"].Value = c.Options;
                workSheet.Range[$"D{i}"].Value = c.Reference;
                workSheet.Range[$"E{i}"].Value = c.Name;
                workSheet.Range[$"F{i}"].Value = c.Value;
                i++;
            }

            // Gets count of fields to assign range for naming workbook commands
            var workbookCommandCount = workbookCommands.Count;
            workbook.CreateNamedRange(nameof(WorkbookCommands), $"B2:B{1 + workbookCommandCount}");
            #endregion

            #region Range Commands
            // Gets commands for each category of methods
            var rangeCommands= CommandHelper.GetRangeCommands();

            // Provides a blank row before starting next set of commands
            i += 1;

            // Gets count of fields to assign range for naming workbook commands
            var rangeCommandCount = rangeCommands.Count;
            workbook.CreateNamedRange(nameof(RangeCommands), $"B{i}:B{i + rangeCommandCount - 1}");

            foreach (var c in rangeCommands)
            {
                workSheet.Range[$"A{i}"].Value = c.CommandType;
                workSheet.Range[$"B{i}"].Value = c.Command;
                workSheet.Range[$"C{i}"].Value = c.Options;
                workSheet.Range[$"D{i}"].Value = c.Reference;
                workSheet.Range[$"E{i}"].Value = c.Name;
                workSheet.Range[$"F{i}"].Value = c.Value;
                i++;
            }
            #endregion

            #region Code Commands
            // Gets commands for each category of methods
            var codeCommands = CommandHelper.GetCodeCommands();

            // Provides a blank row before starting next set of commands
            i += 1;

            // Gets count of fields to assign range for naming workbook commands
            var codeCommandCount = codeCommands.Count;
            workbook.CreateNamedRange(nameof(CodeCommands), $"B{i}:B{i + codeCommandCount - 1}");

            foreach (var c in codeCommands)
            {
                workSheet.Range[$"A{i}"].Value = c.CommandType;
                workSheet.Range[$"B{i}"].Value = c.Command;
                workSheet.Range[$"C{i}"].Value = c.Options;
                workSheet.Range[$"D{i}"].Value = c.Reference;
                workSheet.Range[$"E{i}"].Value = c.Name;
                workSheet.Range[$"F{i}"].Value = c.Value;
                i++;
            }
            #endregion

            #region Styling
            // Styles the command table
            var stylingRange = (Excel.Range)workSheet.Range["A:F"];
            stylingRange.ColumnWidth = 45;
            stylingRange.Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            stylingRange.Cells.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            stylingRange.Cells.WrapText = true;

            // Selects and styles the command headers
            var topRange = workSheet.Range["A1:F1"];
            topRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.CornflowerBlue);
            topRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
            topRange.Font.Bold = true;

            // Alternates color command table rows for easier reading
            for (int o = 3; o < i; o++)
            {
                Excel.Range colorRange;
                if (o%2 != 0)
                {
                    colorRange = workSheet.Range[$"A{o}:F{o}"];
                    colorRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
                    colorRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                }
            }

            // Adds all around border to command table
            Excel.Range borderRange = workSheet.Range[$"A1:F{i-1}"];
            borderRange.Cells.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

            // Adds filter ability to command table for users
            borderRange.AutoFilter(1);

            // Styles the options tables
            stylingRange = (Excel.Range)workSheet.Range["J:L"];
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
            
            XlApp.ActiveWorkbook.CreateNamedRange(workSheet.Name + "Type", "A5");

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

        private void AddWorkbookCommands(object sender, RibbonControlEventArgs e)
        {                    
            Excel.Workbook workbook = Globals.ThisAddIn.Application.ActiveWorkbook;
            Excel.Worksheet worksheet = workbook.ActiveSheet;
            Excel.Range range = Globals.ThisAddIn.Application.ActiveCell;

            //Checks if commands page exists
            if (!workbook.WorksheetExists("Commands"))
            {
                MessageBox.Show("Please run 'Add Commands' and then try again.");
            }

            // Get row and next column indices
            var nextColumn = WorksheetHelper.GetColumnName(range.Column+1);
            var thisRow = range.Row;

            // Command type column
            range.Value = CommandType.Workbook;

            // Command column
            range = worksheet.Range[$"{nextColumn}" + $"{thisRow}"];
            range.AddDropDownList(nameof(WorkbookCommands));

            // Gets next column index
            nextColumn = WorksheetHelper.GetColumnName(range.Column + 1);

            // Options column
            range = worksheet.Range[$"{nextColumn}" + $"{thisRow}"];
            range.AddDropDownList(nameof(WorkbookOptions));

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

        private void AddRangeCommands(object sender, RibbonControlEventArgs e)
        {
            Excel.Workbook workbook = Globals.ThisAddIn.Application.ActiveWorkbook;
            Excel.Worksheet worksheet = workbook.ActiveSheet;
            Excel.Range range = Globals.ThisAddIn.Application.ActiveCell;

            //Checks if commands page exists
            if (!workbook.WorksheetExists("Commands"))
            {
                MessageBox.Show("Please run 'Add Commands' and then try again.");
            }

            // Get row and next column indices
            var nextColumn = WorksheetHelper.GetColumnName(range.Column + 1);
            var thisRow = range.Row;

            // Command type column
            range.Value = CommandType.Range;

            // Command column
            range = worksheet.Range[$"{nextColumn}" + $"{thisRow}"];
            range.AddDropDownList(nameof(RangeCommands));

            // Gets next column index
            nextColumn = WorksheetHelper.GetColumnName(range.Column + 1);

            // Options column
            range = worksheet.Range[$"{nextColumn}" + $"{thisRow}"];
            range.AddDropDownList(nameof(RangeOptions));

            // Gets next column index
            nextColumn = WorksheetHelper.GetColumnName(range.Column + 1);

            // Reference column
            range = worksheet.Range[$"{nextColumn}" + $"{thisRow}"];
            range.AddDropDownList(nameof(ReferenceOptions));
        }

        private void AddCodeCommandsButton_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Workbook workbook = Globals.ThisAddIn.Application.ActiveWorkbook;
            Excel.Worksheet worksheet = workbook.ActiveSheet;
            Excel.Range range = Globals.ThisAddIn.Application.ActiveCell;

            //Checks if commands page exists
            if (!workbook.WorksheetExists("Commands"))
            {
                MessageBox.Show("Please run 'Add Commands' and then try again.");
            }

            // Get row and next column indices
            var nextColumn = WorksheetHelper.GetColumnName(range.Column + 1);
            var thisRow = range.Row;

            // Command type column
            range.Value = CommandType.Code;

            // Command column
            range = worksheet.Range[$"{nextColumn}" + $"{thisRow}"];
            range.AddDropDownList(nameof(CodeCommands));
        }
    }
}
