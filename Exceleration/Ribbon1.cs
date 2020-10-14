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
using System.Collections.Generic;
using Exceleration.DTOS;

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
            CreateCommandWorksheet();
        }

        /// <summary>
        /// Adds command worksheet to workbook. Used for general command explanations and option lists
        /// </summary>
        private void CreateCommandWorksheet()
        {
            // Attempts to find a currently existing Command worksheet
            Excel.Workbook workbook = Globals.ThisAddIn.Application.ActiveWorkbook;
            var worksheet = workbook.GetWorksheets().FirstOrDefault(x => x.Name == "Commands");

            // Adds a worksheet named Commands if it does not already exist
            worksheet = workbook.CreateNewWorksheet("Commands");

            // Adds column headers to command table
            worksheet.Range["A1"].Value = "Command Type";
            worksheet.Range["B1"].Value = "Command";
            worksheet.Range["C1"].Value = "Options";
            worksheet.Range["D1"].Value = "Reference";
            worksheet.Range["E1"].Value = "Name";
            worksheet.Range["F1"].Value = "Value";
            worksheet.Range["G1"].Value = "Output";

            // Add option ranges
            AddOptions(workbook, nameof(WorkbookOptions), "J", OptionHelper.GetWorkbookOptions());
            AddOptions(workbook, nameof(ReferenceOptions), "K", OptionHelper.GetReferenceOptions());
            AddOptions(workbook, nameof(RangeOptions), "L", OptionHelper.GetRangeOptions());

            // Add command ranges
            int counter = 2;
            counter = AddCommands(workbook, nameof(WorkbookCommands), counter, CommandHelper.GetWorkbookCommands());
            counter = AddCommands(workbook, nameof(WorksheetCommands), counter, CommandHelper.GetWorksheetCommands());
            counter = AddCommands(workbook, nameof(RangeCommands), counter, CommandHelper.GetRangeCommands());
            counter = AddCommands(workbook, nameof(CodeCommands), counter, CommandHelper.GetCodeCommands());

            #region Styling
            // Styles the command table
            var stylingRange = (Excel.Range)worksheet.Range["A:G"];
            stylingRange.ColumnWidth = 45;
            stylingRange.Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            stylingRange.Cells.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            stylingRange.Cells.WrapText = true;

            // Selects and styles the command headers
            var topRange = worksheet.Range["A1:G1"];
            topRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.CornflowerBlue);
            topRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
            topRange.Font.Bold = true;

            // Alternates color command table rows for easier reading
            for (int o = 3; o < counter-1; o++)
            {
                Excel.Range colorRange;
                if (o%2 != 0)
                {
                    colorRange = worksheet.Range[$"A{o}:G{o}"];
                    colorRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
                    colorRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                }
            }

            // Adds all around border to command table
            Excel.Range borderRange = worksheet.Range[$"A1:G{counter-2}"];
            borderRange.Cells.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

            // Adds filter ability to command table for users
            borderRange.AutoFilter(1);

            // Styles the options tables
            stylingRange = (Excel.Range)worksheet.Range["J:L"];
            stylingRange.ColumnWidth = 25;
            stylingRange.Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;    
            #endregion
        }

        /// <summary>
        /// Generic method for adding options to the commands page
        /// </summary>
        /// <param name="worksheet">Target worksheet</param>
        /// <param name="name">Name of range being created</param>
        /// <param name="column">Excel column options are to be placed in, alphabetical named not R1C1 style</param>
        /// <param name="options">Selectable options from OptionHelper class</param>
        private void AddOptions(Excel.Workbook workbook, string name, string column, List<string> options)
        {
            Excel.Worksheet worksheet = workbook.ActiveSheet;

            worksheet.Range[$"{column+1}"].Value = name;

            int i = 2;
            foreach(var o in options)
            {
                worksheet.Range[$"{column + i}"].Value = o;
                i++;
            }

            int optionsCount = options.Count;
            workbook.CreateNamedRange(name, $"{column + 2}:{column}{1 + optionsCount}");
        }

        /// <summary>
        /// Generic method for adding commands to the command table
        /// </summary>
        /// <param name="workbook">Target workbook</param>
        /// <param name="commandName">Name of command range</param>
        /// <param name="startingInteger">Excel row index where command is to be added</param>
        /// <param name="commands">List of command items</param>
        /// <returns></returns>
        private int AddCommands(Excel.Workbook workbook, string commandName, int startingInteger,  List<CommandItem> commands)
        {
            Excel.Worksheet worksheet = workbook.ActiveSheet;
            int i = startingInteger;

            foreach (var c in commands)
            {
                worksheet.Range[$"A{i}"].Value = c.CommandType;
                worksheet.Range[$"B{i}"].Value = c.Command;
                worksheet.Range[$"C{i}"].Value = c.Options;
                worksheet.Range[$"D{i}"].Value = c.Reference;
                worksheet.Range[$"E{i}"].Value = c.Name;
                worksheet.Range[$"F{i}"].Value = c.Value;
                worksheet.Range[$"G{i}"].Value = c.Output;
                i++;
            }

            var commandCount = commands.Count;
            workbook.CreateNamedRange(commandName, $"B{startingInteger}:B{startingInteger + commandCount - 1}");

            return i + 1;
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
            workSheet.Range["G5"].Value = "Output";

            XlApp.ActiveWorkbook.CreateNamedRange(workSheet.Name + "Type", "A5");

            // Selects and styles the column headers
            var topRange = workSheet.Range["A5:G5"];
            topRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.CornflowerBlue);
            topRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
            topRange.Font.Bold = true;

            // Styles the command columns
            var stylingRange = (Excel.Range)workSheet.Range["A:G"];
            stylingRange.ColumnWidth = 25;
            stylingRange.Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
        }

        private void AddWorkbookCommands(object sender, RibbonControlEventArgs e)
        {
            AddRowValidation(CommandType.Workbook, nameof(WorkbookCommands), nameof(WorkbookOptions), nameof(ReferenceOptions));
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
            AddRowValidation(CommandType.Range, nameof(RangeCommands), nameof(RangeOptions), nameof(ReferenceOptions));           
        }

        private void AddCodeCommandsButton_Click(object sender, RibbonControlEventArgs e)
        {
            AddRowValidation(CommandType.Code, nameof(CodeCommands));
        }

        private void AddWorksheetCommandsButton_Click(object sender, RibbonControlEventArgs e)
        {
            AddRowValidation(CommandType.Worksheet, nameof(WorksheetCommands));
        }

        /// <summary>
        /// Method for passing in validation to command rows
        /// </summary>
        /// <param name="commandType">Command type value</param>
        /// <param name="command">Command named range</param>
        /// <param name="options">Options named range</param>
        /// <param name="reference">Reference named range</param>
        private void AddRowValidation(string commandType, string command, string options = null, string reference = null)
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
            var thisColumn = WorksheetHelper.GetColumnName(range.Column);
            var thisRow = range.Row;
            Excel.Range startRange = worksheet.Range[$"{thisColumn}" + $"{thisRow}"];

            //Clears any previous validation on the row
            string tempColumn;
            for (int i = 0; i < 6; i++)
            {
                tempColumn = WorksheetHelper.GetColumnName(range.Column + i);
                worksheet.Range[$"{tempColumn}" + $"{thisRow}"].Validation.Delete();
            }

            // Command type column
            range.Value = commandType;

            // Gets next column index
            thisColumn = WorksheetHelper.GetColumnName(range.Column + 1);

            // Command column
            range = worksheet.Range[$"{thisColumn}" + $"{thisRow}"];
            range.AddDropDownList(command);

            // Gets next column index
            thisColumn = WorksheetHelper.GetColumnName(range.Column + 1);

            // Options column
            if (!string.IsNullOrEmpty(options))
            {
                range = worksheet.Range[$"{thisColumn}" + $"{thisRow}"];
                range.AddDropDownList(options);

                // Gets next column index
                thisColumn = WorksheetHelper.GetColumnName(range.Column + 1);


                if(!string.IsNullOrEmpty(reference))
                {
                    // Reference column
                    range = worksheet.Range[$"{thisColumn}" + $"{thisRow}"];
                    range.AddDropDownList(nameof(ReferenceOptions));
                }                
            }            
        }
    }
}
