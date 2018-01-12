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
        public static Form f_Control1;
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            this.button7.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button7_Click);
        }
        private void button7_Click(object sender, RibbonControlEventArgs e)
        {
            f_Control1 = new FormControl1() { TopMost = true };
            f_Control1.Show();
        }
    }
}
