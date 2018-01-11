using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
namespace Excel_Markup
{
    public partial class FormControl1 : Form
    {
        public FormControl1()
        {
            InitializeComponent();
            this.button1.Click += this.button1_Click;
            this.button2.Click += this.button2_Click;
            this.button3.Click += this.button3_Click;
            this.button4.Click += this.button4_Click;
            this.button5.Click += this.button5_Click;
            this.button6.Click += this.button6_Click;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Excel.Range stuff = Globals.ThisAddIn.Application.Selection as Excel.Range;

            stuff.Font.ColorIndex = 3;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Excel.Range stuff = Globals.ThisAddIn.Application.Selection as Excel.Range;

            stuff.Font.ColorIndex = 4;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Excel.Range stuff = Globals.ThisAddIn.Application.Selection as Excel.Range;

            stuff.Font.ColorIndex = 5;
        }
        private void button4_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Excel.Range stuff = Globals.ThisAddIn.Application.Selection as Excel.Range;

            stuff.Font.ColorIndex = 9;
                
        }
        private void button5_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Excel.Range stuff = Globals.ThisAddIn.Application.Selection as Excel.Range;


            stuff.Font.ColorIndex = 7;
        }
        private void button6_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Excel.Range stuff = Globals.ThisAddIn.Application.Selection as Excel.Range;

            dynamic wtf = stuff.Font.Color;

            stuff.Font.ColorIndex = 8;
        }
    }
}
