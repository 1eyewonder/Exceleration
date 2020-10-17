namespace Exceleration
{
    partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon1()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.AddCommandsButton = this.Factory.CreateRibbonButton();
            this.AddTemplateButton = this.Factory.CreateRibbonButton();
            this.RunButton = this.Factory.CreateRibbonButton();
            this.AddWorkbookCommandButton = this.Factory.CreateRibbonButton();
            this.AddWorksheetCommandsButton = this.Factory.CreateRibbonButton();
            this.RangeCommandsButton = this.Factory.CreateRibbonButton();
            this.AddCodeCommandsButton = this.Factory.CreateRibbonButton();
            this.AddFilterCommandsButton = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.group2.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Groups.Add(this.group2);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.AddCommandsButton);
            this.group1.Items.Add(this.AddTemplateButton);
            this.group1.Items.Add(this.RunButton);
            this.group1.Label = "Core";
            this.group1.Name = "group1";
            // 
            // group2
            // 
            this.group2.Items.Add(this.AddWorkbookCommandButton);
            this.group2.Items.Add(this.AddWorksheetCommandsButton);
            this.group2.Items.Add(this.RangeCommandsButton);
            this.group2.Items.Add(this.AddCodeCommandsButton);
            this.group2.Items.Add(this.AddFilterCommandsButton);
            this.group2.Label = "Command Options";
            this.group2.Name = "group2";
            // 
            // AddCommandsButton
            // 
            this.AddCommandsButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.AddCommandsButton.Image = global::Exceleration.Properties.Resources.outline_add_task_black_18dp;
            this.AddCommandsButton.Label = "Add Commands";
            this.AddCommandsButton.Name = "AddCommandsButton";
            this.AddCommandsButton.ShowImage = true;
            this.AddCommandsButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button1_Click);
            // 
            // AddTemplateButton
            // 
            this.AddTemplateButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.AddTemplateButton.Image = global::Exceleration.Properties.Resources.outline_class_black_18dp;
            this.AddTemplateButton.Label = "Add Template";
            this.AddTemplateButton.Name = "AddTemplateButton";
            this.AddTemplateButton.ShowImage = true;
            this.AddTemplateButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.AddTemplateButton_Click);
            // 
            // RunButton
            // 
            this.RunButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.RunButton.Image = global::Exceleration.Properties.Resources.outline_forward_black_18dp;
            this.RunButton.Label = "Run Code";
            this.RunButton.Name = "RunButton";
            this.RunButton.ShowImage = true;
            this.RunButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.RunButton_Click);
            // 
            // AddWorkbookCommandButton
            // 
            this.AddWorkbookCommandButton.Image = global::Exceleration.Properties.Resources.outline_table_view_black_18dp;
            this.AddWorkbookCommandButton.Label = "Add Workbook Commands";
            this.AddWorkbookCommandButton.Name = "AddWorkbookCommandButton";
            this.AddWorkbookCommandButton.ShowImage = true;
            this.AddWorkbookCommandButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.AddWorkbookCommands);
            // 
            // AddWorksheetCommandsButton
            // 
            this.AddWorksheetCommandsButton.Image = global::Exceleration.Properties.Resources.outline_grid_on_black_18dp;
            this.AddWorksheetCommandsButton.Label = "Add Worksheet Commands";
            this.AddWorksheetCommandsButton.Name = "AddWorksheetCommandsButton";
            this.AddWorksheetCommandsButton.ShowImage = true;
            this.AddWorksheetCommandsButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.AddWorksheetCommandsButton_Click);
            // 
            // RangeCommandsButton
            // 
            this.RangeCommandsButton.Image = global::Exceleration.Properties.Resources.outline_view_week_black_18dp;
            this.RangeCommandsButton.Label = "Add Range Commands";
            this.RangeCommandsButton.Name = "RangeCommandsButton";
            this.RangeCommandsButton.ShowImage = true;
            this.RangeCommandsButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.AddRangeCommands);
            // 
            // AddCodeCommandsButton
            // 
            this.AddCodeCommandsButton.Image = global::Exceleration.Properties.Resources.outline_create_black_18dp;
            this.AddCodeCommandsButton.Label = "Add Code Commands";
            this.AddCodeCommandsButton.Name = "AddCodeCommandsButton";
            this.AddCodeCommandsButton.ShowImage = true;
            this.AddCodeCommandsButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.AddCodeCommandsButton_Click);
            // 
            // AddFilterCommandsButton
            // 
            this.AddFilterCommandsButton.Image = global::Exceleration.Properties.Resources.outline_filter_alt_black_18dp;
            this.AddFilterCommandsButton.Label = "Add Filter Commands";
            this.AddFilterCommandsButton.Name = "AddFilterCommandsButton";
            this.AddFilterCommandsButton.ShowImage = true;
            this.AddFilterCommandsButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.AddFilterCommandsButton_Click);
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton AddCommandsButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton AddWorkbookCommandButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton AddTemplateButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton RunButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton RangeCommandsButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton AddCodeCommandsButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton AddWorksheetCommandsButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton AddFilterCommandsButton;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
