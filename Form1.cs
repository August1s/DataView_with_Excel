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
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        ExcelHelper excelhelper = new ExcelHelper(@"E:\C#\DataView_with_Excel\testdata.xlsx");

        private void Form1_Load(object sender, EventArgs e)
        {
            dataGridView1.DataSource = excelhelper.SelectAll();
            //禁止对列的排序
            for (int i = 0; i < dataGridView1.Columns.Count; i++)
            {
                dataGridView1.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
            }
        }

        /// <summary>
        /// 通过用户的id值和id在gridview控件中的位置得到在ecxel中的位置
        /// </summary>
        /// <param name="id">用户的ID</param>
        /// <returns>在excel中的位置</returns>
        private int GetIndex(string id)
        {
            string s;
            for (int i = 0; i < dataGridView1.Rows.Count-1; i++)
            {
                s = dataGridView1.Rows[i].Cells[0].Value.ToString();
                if (id == s)
                    return i+2;
            }
            return -1;
        }

        private void Button_delete_Click(object sender, EventArgs e)
        {
            string userID;

            Form2 f2 = new Form2();
            f2.ShowDialog();
            if (f2.DialogResult == DialogResult.OK)
            {
                userID = f2.ReturnID;
            }
            else
            {
                userID = null;
                return;
            }

            int index = GetIndex(userID);
            if (index == -1)
            {
                MessageBox.Show("wrong userid");
                return;
            }
            excelhelper.DeleteRow(index);
            dataGridView1.DataSource = excelhelper.SelectAll();
        }

        private void Button_insert_Click(object sender, EventArgs e)
        {
            User user;

            Form3 f3 = new Form3();
            f3.ShowDialog();
            if (f3.DialogResult == DialogResult.OK)
            {
                user = f3.ReturnUser;
            }
            else
            {
                user = null;
                return;
            }

            int index = GetIndex(user.UserId);
            if (index != -1)
            {
                MessageBox.Show("userid exist");
                return;
            }
            excelhelper.InsertRow(user);
            dataGridView1.DataSource = excelhelper.SelectAll();
        }

        private void Button_update_Click(object sender, EventArgs e)
        {
            User user;

            Form3 f3 = new Form3();
            f3.ShowDialog();
            if (f3.DialogResult == DialogResult.OK)
            {
                user = f3.ReturnUser;
            }
            else
            {
                user = null;
                return;
            }

            int index = GetIndex(user.UserId);
            if (index == -1)
            {
                MessageBox.Show("wrong userid");
                return;
            }
            excelhelper.UpdateRow(index,user);
            dataGridView1.DataSource = excelhelper.SelectAll();
        }
    }
}
