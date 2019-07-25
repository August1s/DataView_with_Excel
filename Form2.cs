using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace DataView_with_Excel
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
        }


        private string str;
        public string ReturnID
        {
            get { return this.str; }
        }


        private void Button1_Click(object sender, EventArgs e)
        {
            str = this.textBox1.Text;
            
            this.DialogResult = DialogResult.OK;
        }
    }
}
