using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using System.IO;
using System.Reflection;
using Excel = Microsoft.Office.Interop.Excel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;

namespace WindowsFormsApplication2
{
    class Temp
    {

        private void SetOffday()
        {
            /*
            if (this.form_setoffday != null)
                return;

            Form3 form = new Form3(this);

            form.Show();
            form.Focus();
            List<bool> temp = new List<bool>();
            for (int day = 0; day < month_info.Count; day++)
            {
                bool offday = month_info[day].is_offday;
                temp.Add(offday);
            }
            form.ShowCheckBoxs(date_time, temp);
            this.form_setoffday = form;
            */
        }

        private void UpdateExcel()
        {
            /*
            SaveFileDialog dialog = new SaveFileDialog();
            dialog.Filter = "Excel2007|.xlsx";
            dialog.ShowDialog();

            if (dialog.FileName.IndexOf(".xlsx") > 0)
            {
                using (FileStream fs = File.OpenWrite(dialog.FileName))
                {
                    ISheet sheet = workbook.GetSheet("例子");
                    IRow row = sheet.GetRow(1);
                    ICell cell = row.CreateCell(4);
                    cell.SetCellValue("号阿德");

                    workbook.Write(fs);
                }
            }
            */
        }

        bool IsEmptyCell(object cell)
        {
            try
            {
                string value = cell.ToString();
                if (!string.IsNullOrEmpty(value))
                {
                    return false;
                }
                else
                {
                    return true;
                }
            }
            catch (Exception)
            {
                return true;
            }

        }

        private void CreatNewExcel()
        {
            XSSFWorkbook new_excel = new XSSFWorkbook();
            try
            {
                ISheet sheet = new_excel.CreateSheet("例子");
                IRow row = sheet.CreateRow(0);
                ICell cell = row.CreateCell(0);

                DateTime time = DateTime.Now;

                cell.SetCellValue(time);

                SaveFileDialog dialog = new SaveFileDialog();
                dialog.Filter = "Excel2007|.xlsx";
                dialog.ShowDialog();
                if (dialog.FileName.IndexOf(".xlsx") > 0)
                {
                    using (FileStream fs = File.OpenWrite(dialog.FileName))
                    {
                        new_excel.Write(fs);
                    }
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception: " + ex.Message);
            }
        }
    }
}
