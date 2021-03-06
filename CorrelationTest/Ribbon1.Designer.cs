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
            this.GenerateMatrix = this.Factory.CreateRibbonButton();
            this.DeveloperTools = this.Factory.CreateRibbonGroup();
            this.DebugModeToggle = this.Factory.CreateRibbonButton();
            this.testPrint = this.Factory.CreateRibbonButton();
            this.button1 = this.Factory.CreateRibbonButton();
            this.TimingTests = this.Factory.CreateRibbonGroup();
            this.TestDoubles = this.Factory.CreateRibbonButton();
            this.TestStrings = this.Factory.CreateRibbonButton();
            this.TestFormulas = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.group2.SuspendLayout();
            this.group3.SuspendLayout();
            this.DeveloperTools.SuspendLayout();
            this.TimingTests.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Groups.Add(this.group2);
            this.tab1.Groups.Add(this.group3);
            this.tab1.Groups.Add(this.DeveloperTools);
            this.tab1.Groups.Add(this.TimingTests);
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
            this.group3.Items.Add(this.GenerateMatrix);
            this.group3.Label = "Fake Data";
            this.group3.Name = "group3";
            // 
            // FakeFields
            // 
            this.FakeFields.Label = "Fake Fields";
            this.FakeFields.Name = "FakeFields";
            this.FakeFields.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.FakeFields_Click);
            // 
            // GenerateMatrix
            // 
            this.GenerateMatrix.Label = "Test Fit Matrix";
            this.GenerateMatrix.Name = "GenerateMatrix";
            this.GenerateMatrix.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.GenerateMatrix_Click);
            // 
            // DeveloperTools
            // 
            this.DeveloperTools.Items.Add(this.DebugModeToggle);
            this.DeveloperTools.Items.Add(this.testPrint);
            this.DeveloperTools.Items.Add(this.button1);
            this.DeveloperTools.Label = "DeveloperTools";
            this.DeveloperTools.Name = "DeveloperTools";
            // 
            // DebugModeToggle
            // 
            this.DebugModeToggle.Label = "Toggle Debug Mode";
            this.DebugModeToggle.Name = "DebugModeToggle";
            this.DebugModeToggle.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.DebugModeToggle_Click);
            // 
            // testPrint
            // 
            this.testPrint.Label = "Test Print Time";
            this.testPrint.Name = "testPrint";
            this.testPrint.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.testPrint_Click);
            // 
            // button1
            // 
            this.button1.Label = "Test Correl";
            this.button1.Name = "button1";
            this.button1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button1_Click);
            // 
            // TimingTests
            // 
            this.TimingTests.Items.Add(this.TestDoubles);
            this.TimingTests.Items.Add(this.TestStrings);
            this.TimingTests.Items.Add(this.TestFormulas);
            this.TimingTests.Label = "Timing Tests";
            this.TimingTests.Name = "TimingTests";
            // 
            // TestDoubles
            // 
            this.TestDoubles.Label = "Test 1M Doubles";
            this.TestDoubles.Name = "TestDoubles";
            this.TestDoubles.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.TestDoubles_Click);
            // 
            // TestStrings
            // 
            this.TestStrings.Label = "Test 1M Strings";
            this.TestStrings.Name = "TestStrings";
            this.TestStrings.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.TestStrings_Click);
            // 
            // TestFormulas
            // 
            this.TestFormulas.Label = "Test 1M Formulas";
            this.TestFormulas.Name = "TestFormulas";
            this.TestFormulas.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.TestFormulas_Click);
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
            this.DeveloperTools.ResumeLayout(false);
            this.DeveloperTools.PerformLayout();
            this.TimingTests.ResumeLayout(false);
            this.TimingTests.PerformLayout();
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
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup DeveloperTools;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton DebugModeToggle;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton GenerateMatrix;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton testPrint;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup TimingTests;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton TestDoubles;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton TestStrings;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton TestFormulas;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
