using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using System.Data;
using System.Reflection;
using System.Runtime.InteropServices;

namespace DataView_with_Excel
{
    class ExcelHelper
    {
        private string FilePath;


        public ExcelHelper(string path)
        {
            FilePath = path;
        }

        
        [DllImport("User32.dll", CharSet = CharSet.Auto)]
        public static extern int GetWindowThreadProcessId(IntPtr hwnd, out int ID);
        // 由于运行会留下很多excel的进程，手动删除
        private void KillThread(Application excel)
        {
            IntPtr t = new IntPtr(excel.Hwnd);   //得到这个句柄，具体作用是得到这块内存入口 
            int k = 0;
            GetWindowThreadProcessId(t, out k);   //得到本进程唯一标志k 
            System.Diagnostics.Process p = System.Diagnostics.Process.GetProcessById(k);   //得到对进程k的引用 
            p.Kill();     //关闭进程k 

        }


        /// <summary>
        /// 从文件路径中读取excel的值
        /// </summary>
        /// <returns>返回一个datatable类</returns>
        public System.Data.DataTable SelectAll()
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            ApplicationClass excelapp = new ApplicationClass();
            Workbook workbook = excelapp.Workbooks.Open(FilePath, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            Worksheet worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Worksheets.get_Item(1);

            if (worksheet == null)
                return null;

            int iRowCount = worksheet.UsedRange.Rows.Count;
            int iColCount = worksheet.UsedRange.Columns.Count;
            //生成列
            for (int i = 0; i < iColCount; i++)
            {
                var name = "column" + i;
                var txt = ((Range)worksheet.Cells[1, i + 1]).Text.ToString();
                
                if (!string.IsNullOrWhiteSpace(txt)) name = txt;//防止列为空
                while (dt.Columns.Contains(name)) name = name + "_1";//防止重复列名称。
                dt.Columns.Add(new DataColumn(name, typeof(string)));
            }
            //生成行
            Range range;
            int rowIdx = 2;
            for (int iRow = rowIdx; iRow <= iRowCount; iRow++)
            {
                DataRow dr = dt.NewRow();
                for (int iCol = 1; iCol <= iColCount; iCol++)
                {
                    range = (Range)worksheet.Cells[iRow, iCol];
                    dr[iCol - 1] = (range.Value2 == null) ? "" : range.Text.ToString();
                }
                dt.Rows.Add(dr);
            }
            
            workbook.Close(Missing.Value, Missing.Value, Missing.Value);
            excelapp.Workbooks.Close();
            excelapp.Quit();
            KillThread(excelapp);
            return dt;
        }


        /// <summary>
        /// 删除一行
        /// </summary>
        /// <param name="rowIndex">行索引</param>
        public void DeleteRow(int rowIndex)
        {
            ApplicationClass excelapp = new ApplicationClass();
            //excelapp.Visible = true;

            Workbook workbook = excelapp.Workbooks.Open(FilePath, Type.Missing, false, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            Worksheet worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Worksheets.get_Item(1);

            Range range = (Range)worksheet.Rows[rowIndex,Missing.Value];    //获取修改范围
            range.Delete(XlDirection.xlDown);

            //workbook.SaveAs(FilePath, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            workbook.Close(true, Missing.Value, Missing.Value);
            excelapp.Workbooks.Close();
            excelapp.Quit();
            KillThread(excelapp);

        }


        /// <summary>
        /// 按照用户id更新一行
        /// </summary>
        /// <param name="rowIndex">行索引</param>
        /// <param name="user">用户实例</param>
        public void UpdateRow(int rowIndex,User user )
        {
            ApplicationClass excelapp = new ApplicationClass();
            //excelapp.Visible = true;

            Workbook workbook = excelapp.Workbooks.Open(FilePath, Type.Missing, false, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            Worksheet worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Worksheets.get_Item(1);

            worksheet.Cells[rowIndex, 1] = user.UserId;
            worksheet.Cells[rowIndex, 2] = user.UserName;
            worksheet.Cells[rowIndex, 3] = user.UserAge;

            //workbook.SaveAs(FilePath, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            workbook.Close(true, Missing.Value, Missing.Value);
            excelapp.Workbooks.Close();
            excelapp.Quit();
            KillThread(excelapp);

        }


        /// <summary>
        /// 在表格末端插入一行
        /// </summary>
        /// <param name="user">用户实例</param>
        public void InsertRow(User user)
        {
            ApplicationClass excelapp = new ApplicationClass();
            //excelapp.Visible = true;

            Workbook workbook = excelapp.Workbooks.Open(FilePath, Type.Missing, false, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            Worksheet worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Worksheets.get_Item(1);
            int iRowCount = worksheet.UsedRange.Rows.Count;

            worksheet.Cells[iRowCount + 1, 1] = user.UserId;
            worksheet.Cells[iRowCount + 1, 2] = user.UserName;
            worksheet.Cells[iRowCount + 1, 3] = user.UserAge;

            //workbook.SaveAs(FilePath, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            workbook.Close(true, Missing.Value, Missing.Value);
            excelapp.Workbooks.Close();
            excelapp.Quit();
            KillThread(excelapp);

        }
    }
}
