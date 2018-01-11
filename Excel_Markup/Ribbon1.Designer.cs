namespace Excel_Markup
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
            this.button_Annual = this.Factory.CreateRibbonButton();
            this.button_Fiscal = this.Factory.CreateRibbonButton();
            this.button_SchoolYear = this.Factory.CreateRibbonButton();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.button_Series = this.Factory.CreateRibbonButton();
            this.group3 = this.Factory.CreateRibbonGroup();
            this.button_Category = this.Factory.CreateRibbonButton();
            this.button_OtherCategories = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.group2.SuspendLayout();
            this.group3.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.Groups.Add(this.group1);
            this.tab1.Groups.Add(this.group2);
            this.tab1.Groups.Add(this.group3);
            this.tab1.Label = "Excel Markup";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.button_Annual);
            this.group1.Items.Add(this.button_Fiscal);
            this.group1.Items.Add(this.button_SchoolYear);
            this.group1.Label = "Date";
            this.group1.Name = "group1";
            // 
            // button_Annual
            // 
            this.button_Annual.Label = "Annual";
            this.button_Annual.Name = "button_Annual";
            // 
            // button_Fiscal
            // 
            this.button_Fiscal.Label = "Fiscal";
            this.button_Fiscal.Name = "button_Fiscal";
            // 
            // button_SchoolYear
            // 
            this.button_SchoolYear.Label = "SchoolYear";
            this.button_SchoolYear.Name = "button_SchoolYear";
            // 
            // group2
            // 
            this.group2.Items.Add(this.button_Series);
            this.group2.Label = "group2";
            this.group2.Name = "group2";
            // 
            // button_Series
            // 
            this.button_Series.Label = "Series";
            this.button_Series.Name = "button_Series";
            // 
            // group3
            // 
            this.group3.Items.Add(this.button_Category);
            this.group3.Items.Add(this.button_OtherCategories);
            this.group3.Label = "group3";
            this.group3.Name = "group3";
            // 
            // button_Category
            // 
            this.button_Category.Label = "Category";
            this.button_Category.Name = "button_Category";
            // 
            // button_OtherCategories
            // 
            this.button_OtherCategories.Label = "Other Categories";
            this.button_OtherCategories.Name = "button_OtherCategories";
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
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button_Annual;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button_Fiscal;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button_SchoolYear;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button_Series;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button_Category;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button_OtherCategories;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
