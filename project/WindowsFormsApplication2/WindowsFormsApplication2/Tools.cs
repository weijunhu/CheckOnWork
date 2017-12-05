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
    public class Tools
    {
         
        public static DateTime GetDateFromeCell(ICell cell_day)
        {
            DateTime time_temp = DateTime.Now;
            if (cell_day.CellType == CellType.String)
            {
                string[] s = cell_day.StringCellValue.Split('/');
                int year = Convert.ToInt32(s[0]);
                int month = Convert.ToInt32(s[1]);
                int day = Convert.ToInt32(s[2]);
                time_temp = new DateTime(year, month, day);
            }
            else if (cell_day.CellType == CellType.Numeric)
                time_temp = cell_day.DateCellValue;
            else
                time_temp = new System.DateTime(2000, 1, 1);

            return time_temp;
        }

        public static DateTime GetTimeFromeCell(ICell cell_time)
        {
            DateTime temp;
            if (cell_time.CellType == CellType.Blank)
                temp = new DateTime(2000, 1, 1, 0, 0, 0);
            else if (cell_time.CellType == CellType.String)
            {
                if (string.IsNullOrEmpty(cell_time.StringCellValue))
                {
                    temp = new DateTime(2000, 1, 1, 0, 0, 0);
                }
                else
                {
                    string[] s = cell_time.StringCellValue.Split(':');
                    int h = Convert.ToInt32(s[0]);
                    int m = Convert.ToInt32(s[1]);
                    temp = new DateTime(2000, 1, 1, h, m, 0);
                }
            }
            else
                temp = cell_time.DateCellValue;

            return temp;
        }

        public static bool IsLater()
        {


            return false;
        }

        public static bool IsLater(DayInfo day_info)
        {
            if (IsOffday(day_info.date.Day))
                return false;

            if (day_info.no_sign_in)
                return false;
            DateTime time = day_info.time_sign_in;
            TimeSpan start = new TimeSpan(time.Hour, time.Minute, time.Second);
            TimeSpan span = new TimeSpan(9, 35, 0);
            TimeSpan span_tmp = new TimeSpan(10, 5, 0);
            if (day_info.date.Day == 1)
                return start > span;
            else
            {
                Console.WriteLine("<<<");
                PlayerInfo player_info = AllPlayer.all_player_info[day_info.name];
                DayInfo info = player_info.GetDayInfo(day_info.date.Day - 1);//前一天
                bool t = info.is_workday && info.is_allowance;
                if (t)
                    return start > span_tmp;
                else
                    return start > span;
            }
        }

        public static bool IsOffday(int day)
        {
            return MonthInfo.month_info[day - 1].is_offday;
        }

        public static bool IsDeductionLater(DayInfo day_info)
        {
            if (!day_info.is_later)
                return false;
            if (IsOffday(day_info.date.Day))
                return false;
            if (day_info.no_sign_in)
                return false;
            DateTime time = day_info.time_sign_in;
            TimeSpan start = new TimeSpan(time.Hour, time.Minute, time.Second);
            TimeSpan span = new TimeSpan(9, 35, 0);
            TimeSpan span_tmp = new TimeSpan(10, 5, 0);
            if (day_info.date.Day == 1)
                return false;
            else
            {
                Console.WriteLine(">>>");
                PlayerInfo player_info = AllPlayer.all_player_info[day_info.name];
                DayInfo info = player_info.GetDayInfo(day_info.date.Day - 1);//前一天
                bool t = info.is_workday && info.is_allowance;
                if (t)
                    return start >= span && start <= span_tmp;
                else
                    return false;
            }
        }

        public static bool IsAllowance(DayInfo day_info)
        {
            if (day_info.no_sign_out)
                return false;
            if (IsOffday(day_info.date.Day))
                return true;

            DateTime time = day_info.time_sign_out;
            TimeSpan end = new TimeSpan(time.Hour, time.Minute, time.Second);
            TimeSpan span = new TimeSpan(21, 0, 0);
            return end >= span ;
        }
   
    }
}
