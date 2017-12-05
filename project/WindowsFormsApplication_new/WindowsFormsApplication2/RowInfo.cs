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
    public class RowInfo
    {
        public string name;
        public DateTime start;
        public DateTime end;
        public DateTime date;
        public bool is_no_start;
        public bool is_no_end;

        public RowInfo(IRow row)
        {
            ICell cell_name = row.GetCell(0);
            ICell cell_day = row.GetCell(1);
            ICell cell_start = row.GetCell(2);
            ICell cell_end = row.GetCell(3);


            name = cell_name.StringCellValue;

            if (cell_start.CellType == CellType.Blank)
            {
                start = new DateTime(2000, 1, 1, 0, 0, 0);
                is_no_start = true;
            }
            else if (cell_start.CellType == CellType.String)
            {
                if (string.IsNullOrEmpty(cell_start.StringCellValue))
                {
                    start = new DateTime(2000, 1, 1, 0, 0, 0);
                    is_no_start = true;
                }
                else
                {
                    string[] s = cell_start.StringCellValue.Split(':');
                    int h = Convert.ToInt32(s[0]);
                    int m = Convert.ToInt32(s[1]);
                    start = new DateTime(2000, 1, 1, h, m, 0);
                }
            }
            else
                start = cell_start.DateCellValue;

            if (cell_end.CellType == CellType.Blank)
            {
                end = new DateTime(2000, 1, 1, 0, 0, 0);
                is_no_end = true;
            }  
            else if (cell_end.CellType == CellType.String)
            {
                if (string.IsNullOrEmpty(cell_end.StringCellValue))
                {
                    end = new DateTime(2000, 1, 1, 0, 0, 0);
                    is_no_end = true;
                }
                else
                {
                    Console.WriteLine(cell_end.StringCellValue);
                    string[] s = cell_end.StringCellValue.Split(':');
                    int h = Convert.ToInt32(s[0]);
                    int m = Convert.ToInt32(s[1]);
                    end = new DateTime(2000, 1, 1, h, m, 0);
                }
            }
            else
                end = cell_end.DateCellValue;


            if (cell_day.CellType == CellType.String)
            {
                string[] s = cell_day.StringCellValue.Split('/');
                int year = Convert.ToInt32(s[0]);
                int month = Convert.ToInt32(s[1]);
                int day = Convert.ToInt32(s[2]);
                date = new DateTime(year, month, day);
            }
            else if (cell_day.CellType == CellType.Numeric)
            {
                date = cell_day.DateCellValue;
            }
        }
    }
}
