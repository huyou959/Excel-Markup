using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using System.Windows.Forms;
using System.IO;
using System.Diagnostics;
using System.Reflection;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;

namespace Excel_Markup
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            this.button_Annual.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button_Annual_Click);
            this.button_Fiscal.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button_Fiscal_Click);
            this.button_SchoolYear.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button_SchoolYear_Click);
            this.button_Series.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button_Series_Click);
            this.button_Category.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button_Category_Click);
            this.button_OtherCategories.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button_OtherCategories_Click);

        }
        private void button_Annual_Click(object sender, RibbonControlEventArgs e)
        {   
            MessageBox.Show("You are clicking Annual Button!");

            Microsoft.Office.Interop.Excel.Range stuff = Globals.ThisAddIn.Application.Selection as Excel.Range;

          dynamic wtf=  stuff.Font.Color;

            stuff.Font.ColorIndex = 4;
            // write the main function here
        }
        private void button_Fiscal_Click(object sender, RibbonControlEventArgs e)
        {
            MessageBox.Show("You are clicking Fiscal Button!");
        }
        private void button_SchoolYear_Click(object sender, RibbonControlEventArgs e)
        {
            MessageBox.Show("You are clicking SchoolYear Button!");
        }
        private void button_Series_Click(object sender, RibbonControlEventArgs e)
        {
            MessageBox.Show("You are clicking Series Button!");
        }
        private void button_Category_Click(object sender, RibbonControlEventArgs e)
        {
            MessageBox.Show("You are clicking Category Button!");
        }
        private void button_OtherCategories_Click(object sender, RibbonControlEventArgs e)
        {
            MessageBox.Show("You are clicking OtherCategories Button!");
        }
    }
}
