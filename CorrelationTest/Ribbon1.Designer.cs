namespace CorrelationTest
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
            this.BuildCorrelation = this.Factory.CreateRibbonButton();
            this.ExpandCorrel = this.Factory.CreateRibbonButton();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.CollapseCorrel = this.Factory.CreateRibbonButton();
            this.btnVisualize = this.Factory.CreateRibbonButton();
            this.group3 = this.Factory.CreateRibbonGroup();
            this.FakeFields = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.group2.SuspendLayout();
            this.group3.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Groups.Add(this.group2);
            this.tab1.Groups.Add(this.group3);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.BuildCorrelation);
            this.group1.Items.Add(this.ExpandCorrel);
            this.group1.Label = "CorrelString";
            this.group1.Name = "group1";
            // 
            // BuildCorrelation
            // 
            this.BuildCorrelation.Label = "BuildCorrelation";
            this.BuildCorrelation.Name = "BuildCorrelation";
            this.BuildCorrelation.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BuildCorrelation_Click);
            // 
            // ExpandCorrel
            // 
            this.ExpandCorrel.Label = "Expand Correlation";
            this.ExpandCorrel.Name = "ExpandCorrel";
            this.ExpandCorrel.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ExpandCorrel_Click);
            // 
            // group2
            // 
            this.group2.Items.Add(this.CollapseCorrel);
            this.group2.Items.Add(this.btnVisualize);
            this.group2.Label = "CorrelSheet";
            this.group2.Name = "group2";
            // 
            // CollapseCorrel
            // 
            this.CollapseCorrel.Label = "Collapse Correlation";
            this.CollapseCorrel.Name = "CollapseCorrel";
            this.CollapseCorrel.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.CollapseCorrel_Click);
            // 
            // btnVisualize
            // 
            this.btnVisualize.Label = "Visualize";
            this.btnVisualize.Name = "btnVisualize";
            this.btnVisualize.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnVisualize_Click);
            // 
            // group3
            // 
            this.group3.Items.Add(this.FakeFields);
            this.group3.Label = "Fake Data";
            this.group3.Name = "group3";
            // 
            // FakeFields
            // 
            this.FakeFields.Label = "Fake Fields";
            this.FakeFields.Name = "FakeFields";
            this.FakeFields.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.FakeFields_Click);
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
            this.group3.ResumeLayout(false);
            this.group3.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BuildCorrelation;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ExpandCorrel;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton CollapseCorrel;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton FakeFields;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnVisualize;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
