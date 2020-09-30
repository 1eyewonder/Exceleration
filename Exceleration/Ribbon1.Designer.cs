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
            this.AddSheetCommandButton = this.Factory.CreateRibbonButton();
            this.AddTemplateButton = this.Factory.CreateRibbonButton();
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
            this.group1.Label = "Core";
            this.group1.Name = "group1";
            // 
            // group2
            // 
            this.group2.Items.Add(this.AddSheetCommandButton);
            this.group2.Label = "Command Options";
            this.group2.Name = "group2";
            // 
            // AddCommandsButton
            // 
            this.AddCommandsButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.AddCommandsButton.Label = "Add Commands";
            this.AddCommandsButton.Name = "AddCommandsButton";
            this.AddCommandsButton.ShowImage = true;
            this.AddCommandsButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button1_Click);
            // 
            // AddSheetCommandButton
            // 
            this.AddSheetCommandButton.Label = "Add Sheet Commands";
            this.AddSheetCommandButton.Name = "AddSheetCommandButton";
            this.AddSheetCommandButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.AddSheetCommands);
            // 
            // AddTemplateButton
            // 
            this.AddTemplateButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.AddTemplateButton.Label = "Add Template";
            this.AddTemplateButton.Name = "AddTemplateButton";
            this.AddTemplateButton.ShowImage = true;
            this.AddTemplateButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.AddTemplateButton_Click);
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
        internal Microsoft.Office.Tools.Ribbon.RibbonButton AddSheetCommandButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton AddTemplateButton;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
