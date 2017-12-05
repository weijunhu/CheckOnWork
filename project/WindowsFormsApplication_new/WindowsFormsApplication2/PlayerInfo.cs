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
    public class PlayerInfo
    {
        public string name;

        Dictionary<int, DayInfo> date_info = new Dictionary<int, DayInfo>();

        public void AddDayInfo(IRow row)
        {
            RowInfo row_info = new RowInfo(row);
            DayInfo day_info = new DayInfo(row_info);
            name = row_info.name;
            int day = row_info.date.Day;
            date_info.Add(day, day_info);
        }

        public void AddDayInfo(DayInfo day_info)
        {
            name = day_info.name;
            int day = day_info.date.Day;
            date_info.Add(day, day_info);
        }

        public DayInfo GetDayInfo(int day)
        {
            Console.WriteLine(day);
            return date_info[day];
        }

        public int GetCountLater()
        {
            int count = 0;
            foreach (KeyValuePair<int, DayInfo> kvp in date_info)
            {
                if (kvp.Value.is_later)
                    count += 1;
            }
            return count;
        }

        public int GetCountNotsign()
        {
            int count = 0;
            foreach (KeyValuePair<int, DayInfo> kvp in date_info)
            {
                if (kvp.Value.no_sign_in || kvp.Value.no_sign_out)
                    count += 1;
            }
            return count;
        }

        public int GetCountAllowance()
        {
            int count = 0;
            foreach (KeyValuePair<int, DayInfo> kvp in date_info)
            {
                if (kvp.Value.is_allowance)
                    count += 1;
            }
            return count;
        }

    }
}
