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
    public partial class Form3 : Form
    {
        public Form3()
        {
            InitializeComponent();
        }

        private User user = new User();
        public User ReturnUser
        {
            get { return this.user; }   
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            user.UserId =  textBox1.Text;
            user.UserName = textBox2.Text;
            user.UserAge = textBox3.Text;

            this.DialogResult = DialogResult.OK;
        }
    }
}
