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
    public class DayInfo
    {
        public string name;
        public DateTime date;
        public DateTime time_sign_in;
        public DateTime time_sign_out;

        public bool no_sign_in;
        public bool no_sign_out;
        public bool is_later;
        public bool is_workday;
        public bool is_offday;
        public bool is_deduction_later;
        public bool is_allowance;

        RowInfo row;

        public DayInfo(RowInfo row_info)
        {
            row = row_info;

            name = row_info.name;
            date = row_info.date;

            time_sign_in = row_info.start;
            time_sign_out = row_info.end;

            no_sign_in = row_info.is_no_start;
            no_sign_out = row_info.is_no_end;

            is_later = Tools.IsLater(this);
            is_offday = Tools.IsOffday(row_info.date.Day);
            is_workday = !is_offday;
            is_deduction_later = Tools.IsDeductionLater(this);
            is_allowance = Tools.IsAllowance(this);
        }
    }
}
