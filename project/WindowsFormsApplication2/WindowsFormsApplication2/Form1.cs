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
    public partial class Form1 : Form
    {
        Form form_setoffday;
        private DateTime date_time;
        /// <summary>
        /// 所有数据
        /// </summary>
        public Dictionary<string, player_info> dic_all_player = new Dictionary<string, player_info>();

        private IWorkbook workbook = null;
        private ISheet cur_sheet = null;

        string[] weekdays = { "星期日", "星期一", "星期二", "星期三", "星期四", "星期五", "星期六" };

        /// <summary>
        /// 人员数据项
        /// </summary>
        public class player_info
        {
            public struct time_info
            {
                public DateTime date;
                public DateTime time_start;
                public DateTime time_end;
                public bool is_allowance;
                public bool is_workday;
                public bool is_later;
                public bool is_deduction_later;
                public int is_not_sing;
            }
            public string name;
            public int count_later;
            public int count_notsign;
            public int count_allowance;
            public int count_later_in_five;
            public int count_overtime;
            public List<time_info> date_info = new List<time_info>();
        }

        

        public Form1()
        {
               InitializeComponent();       
        }

       



        private void InitOrgDate()
        {
            OpenFileDialog open_file = new OpenFileDialog();
            open_file.Filter = "Excel|*.xlsx;*.xls";
            open_file.ShowDialog();

            string path = open_file.FileName;
            this.label1.Text = string.Format("Excel位置：{0}", path);

            dic_all_player.Clear();
            this.listView1.Clear();

            if (string.IsNullOrEmpty(path))
                return;

            using (FileStream fs = File.OpenRead(path))
            {
                if (path.IndexOf(".xlsx") > 0) // 2007版本
                    workbook = new XSSFWorkbook(fs);
                else if (path.IndexOf(".xls") > 0) // 2003版本we
                    workbook = new HSSFWorkbook(fs);

                ISheet sheet = workbook.GetSheetAt(0);
                if (sheet == null)
                {
                    MessageBox.Show("当前是个空表！");
                    return;
                }

                cur_sheet = sheet;

                IRow row = sheet.GetRow(1);
                ICell cell_name = row.GetCell(0);
                ICell cell_day = row.GetCell(1);

                if (cell_name.CellType == CellType.Blank)
                {
                    MessageBox.Show("请检查表！");
                    return;
                } 

                date_time = Tools.GetDateFromeCell(cell_day);

                String str = String.Format("检测到你要统计的是{0}年{1}月份的考勤，是否继续？",date_time.Year, date_time.Month);
                DialogResult dialog = MessageBox.Show(str, "赶紧回答！", MessageBoxButtons.YesNo);

                if (dialog == DialogResult.Yes)
                {
                    InitDayInfo();
                    LoadExcelData();
                    //SetOffday();
                }
                else
                {
                    this.Close();
                }
            }
        }

        private void InitDayInfo()
        {
            int month = date_time.Month;

            DateTime date = new DateTime(date_time.Year, month, 1);
            int day_count = DateTime.DaysInMonth(date_time.Year, month);

            for (int i = 0; i < day_count; i++)
            {
                int day = i + 1;

                DateTime date_temp = new DateTime(date_time.Year, month, day);
                DayOfWeek week = date_temp.DayOfWeek;
                Day info = new Day();
                info.date = date_temp;
                //定义休息日
                info.is_offday = week == DayOfWeek.Sunday || week == DayOfWeek.Saturday;
                MonthInfo.month_info.Add(info);
            }
        }


        private void LoadExcelData()
        {

            ISheet sheet = cur_sheet;
            int count_row = sheet.LastRowNum;

            //Console.WriteLine(count_row);
            this.progressBar1.Show();
            this.progressBar1.Minimum = 0;
            this.progressBar1.Maximum = count_row;
            for (int i = 1; i < count_row; i++)
            {
                this.label2.Text = string.Format("正在加载第{0}条数据", i);
                IRow row = sheet.GetRow(i);

                RowInfo row_info = new RowInfo(row);
               
                InitPlayerInfo(row_info);

                this.progressBar1.Value = i;
            }
            this.progressBar1.Hide();
            this.label2.Text = "";
            UpdateList();
        }


        private void UpdateList()
        {
            int[] width = { 50,40, 85, 85 };
            string[] tip = { "姓名","序号", "迟到", "未打卡" };
            for (int i = 0; i < tip.Length; i++)
            {
                ColumnHeader ch = new ColumnHeader();
                ch.Text = tip[i];
                ch.Width = width[i];
                ch.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
                this.listView1.Columns.Add(ch);
            }

            this.listView1.BeginUpdate();

            ListViewGroup nomal = new ListViewGroup();
            nomal.Header = "正常打卡";
            nomal.HeaderAlignment = System.Windows.Forms.HorizontalAlignment.Left;

            ListViewGroup unusual = new ListViewGroup();
            unusual.Header = "打卡异常";
            unusual.HeaderAlignment = System.Windows.Forms.HorizontalAlignment.Left;

            //this.listView1.Groups.Add(nomal);
            //this.listView1.Groups.Add(unusual);
            //this.listView1.ShowGroups = true;

            int count = 1;

            foreach (KeyValuePair<string, PlayerInfo> kvp in AllPlayer.all_player_info)
            {
                ListViewItem lvi = new ListViewItem();
                //lvi.ImageIndex = i;
                lvi.Text = kvp.Key;
                lvi.SubItems.Add(count.ToString());
                int c_later = kvp.Value.GetCountLater();
                int c_sign = kvp.Value.GetCountNotsign();
                lvi.SubItems.Add(c_later>0?c_later + "次":"");
                lvi.SubItems.Add(c_sign > 0 ? c_sign + "次" :"");
                /*
                if (kvp.Value.count_later > 0)
                {
                    lvi.ForeColor = Color.Red;
                    unusual.Items.Add(lvi);
                }
                else
                    nomal.Items.Add(lvi);
                */
                this.listView1.Items.Add(lvi);
                count += 1;
            }
            this.listView1.EndUpdate();
        }


       // private void InitPlayerInfo(string name, DateTime time_start, DateTime time_end, int day)
        private void InitPlayerInfo(RowInfo row_info)
        {
            DayInfo day_info = new DayInfo(row_info);

            PlayerInfo player_info;

            if (AllPlayer.all_player_info.ContainsKey(row_info.name))
                player_info = AllPlayer.all_player_info[row_info.name];
            else
            {
                player_info = new PlayerInfo();
                player_info.AddDayInfo(day_info);
                AllPlayer.all_player_info.Add(row_info.name, player_info);
            }
        }



        private bool is_offday(int day)
        {
            return MonthInfo.month_info[day - 1].is_offday;
        }

        public string Week(DateTime time)
        {
            string week = weekdays[Convert.ToInt32(time.DayOfWeek)];
            return week;
        }

        /// <summary>
        ///     
        /// </summary>
        /// <param name="time"></param>
        /// <param name="name"></param>
        /// <param name="index"></param>
        /// <returns></returns>
        public bool IsDeduction(DateTime time, string name, int index, DateTime date)
        {
            if (is_offday(time.Day))
                return false;

            TimeSpan start = new TimeSpan(time.Hour, time.Minute, time.Second);
            TimeSpan span = new TimeSpan(9, 35, 0);
            TimeSpan span_tmp = new TimeSpan(10, 0, 5);
            if (index == 0)
                return false;
            else
            {
                player_info info = dic_all_player[name];
                bool t = info.date_info[index - 1].is_workday && info.date_info[index - 1].is_allowance;
                if (t)
                    return start >= span && start <= span_tmp;
                else
                    return false;
            }
        }

        public bool IsLater(DateTime time, string name, int index, DateTime date)
        {
            if (is_offday(date.Day))
                return false;

            TimeSpan start = new TimeSpan(time.Hour, time.Minute, time.Second);
            TimeSpan span = new TimeSpan(9, 35, 0);
            TimeSpan span_tmp = new TimeSpan(10, 0, 5);
            if (index == 0)
                return start > span;
            else
            {
                player_info info = dic_all_player[name];
                bool t = info.date_info[index - 1].is_workday && info.date_info[index - 1].is_allowance;
                if (t)
                    return start > span_tmp;
                else
                    return start > span;
            }
        }

        private bool IsAllowance(TimeSpan time_end,TimeSpan time_start, DateTime date)
        {
            TimeSpan allowance = new TimeSpan(22, 0, 0);
            if (is_offday(date.Day) && (time_end != TimeSpan.Zero || time_start != TimeSpan.Zero))
                return true;
            else if (time_end != TimeSpan.Zero && time_end >= allowance)
                return true;
            else
                return false;
        }

        private int IsNotSign(TimeSpan time_end, TimeSpan time_start, DateTime date)
        {
            if (is_offday(date.Day))
                return 0;
            else if (time_end == TimeSpan.Zero && time_start == TimeSpan.Zero)
                return 2;
            else if (time_end == TimeSpan.Zero && time_start != TimeSpan.Zero)
                return 1;
            else if (time_end != TimeSpan.Zero && time_start == TimeSpan.Zero)
                return 1;
            else
                return 0;
        }


        private void InitPlayerDetail(player_info info)
        {
            this.listView2.Clear();
            int[] width = { 30,50, 80, 80, 50, 60, 60, 60 };
            string[] tip = { "天","星期几", "上班", "下班" ,"休息","打卡","餐补","迟到" };
            for (int i = 0; i < 8; i++)
            {
                ColumnHeader ch = new ColumnHeader();
                ch.Text = tip[i];
                ch.Width = width[i];
                ch.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
                this.listView2.Columns.Add(ch);
            }

            this.listView2.BeginUpdate();

            int count_allowance = 0;
            int count_later = 0;
            int count_not_sign = 0;
            for (int i = 0; i < info.date_info.Count; i++ )
            {
                player_info.time_info t_info = info.date_info[i];

                ListViewItem lvi = new ListViewItem();
                //lvi.ImageIndex = i;
                lvi.Text = (i+1).ToString();

                lvi.SubItems.Add(Week(t_info.date));


                TimeSpan start = new TimeSpan(t_info.time_start.Hour, t_info.time_start.Minute, t_info.time_start.Second);
                string str_start = "";

                //--1
                if (start != TimeSpan.Zero)
                    str_start = start.ToString();
                lvi.SubItems.Add(str_start);

                //--2
                TimeSpan end = new TimeSpan(t_info.time_end.Hour, t_info.time_end.Minute, t_info.time_end.Second);
                string str_end = "";
                if (end != TimeSpan.Zero)
                    str_end = end.ToString();
                lvi.SubItems.Add(str_end);

                //--3
                if (is_offday(t_info.date.Day))
                    lvi.SubItems.Add("休息日");
                else
                    lvi.SubItems.Add("");

                //--4
                if (t_info.is_not_sing > 0)
                {
                    count_not_sign += t_info.is_not_sing;
                    lvi.SubItems.Add("卡x " + count_not_sign);
                }   
                else
                    lvi.SubItems.Add("");

                //--5
                if (t_info.is_allowance)
                {
                    count_allowance += 1;
                    lvi.SubItems.Add("补x " + count_allowance);
                }   
                else
                    lvi.SubItems.Add("");

                //--6
                if (t_info.is_later)
                {
                    count_later += 1;
                    lvi.SubItems.Add("迟x " + count_later);
                } 
                else if (t_info.is_deduction_later)
                    lvi.SubItems.Add("抵扣");
                else
                    lvi.SubItems.Add("");

                this.listView2.Items.Add(lvi);
            }
            this.listView2.EndUpdate();
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void listView1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (this.listView1.SelectedItems.Count > 0)
            {
                this.groupBox1.Text = this.listView1.SelectedItems[0].Text;

                player_info info = dic_all_player[this.listView1.SelectedItems[0].Text];
                InitPlayerDetail(info);
            }
        }

        private void listView2_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (this.listView2.SelectedItems.Count > 0)
            {
                //this.groupBox3.Text = this.listView1.SelectedItems[0].Text;
            }
        }

        private void 导入ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //CreatNewExcel();
            //UpdateList();
            //OpenExcel();
            //UpdateExcel();
            InitOrgDate();
        }



        private void 导出ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            CreatNewSheet();
        }


        private void CreatNewSheet()
        {
            SaveFileDialog dialog = new SaveFileDialog();
            dialog.Filter = "Excel2007|.xlsx";
            dialog.ShowDialog();

            if (dialog.FileName.IndexOf(".xlsx") > 0)
            {
                using (FileStream fs = File.OpenWrite(dialog.FileName))
                {
                    ISheet sheet = workbook.CreateSheet("合并");
                    IRow row = sheet.CreateRow(0);
                    ICell cell_1 = row.CreateCell(0);
                    cell_1.SetCellValue("姓名");

                    ICell cell_2 = row.CreateCell(1);
                    cell_2.SetCellValue("迟到");

                    ICell cell_3 = row.CreateCell(2);
                    cell_3.SetCellValue("没打卡");

                    ICell cell_4 = row.CreateCell(3);
                    cell_4.SetCellValue("餐补");

                    WriteCombinedSheet(sheet);

                    workbook.Write(fs);
                }
            }
        }

        private void WriteCombinedSheet(ISheet sheet)
        {
            int index = 1;
            foreach (KeyValuePair<string, player_info> kvp in dic_all_player)
            {
                IRow row = sheet.CreateRow(index);
                ICell cell_1 = row.CreateCell(0);
                cell_1.SetCellValue(kvp.Key);

                ICell cell_2 = row.CreateCell(1);
                cell_2.SetCellValue(kvp.Value.count_later);

                ICell cell_3 = row.CreateCell(2);
                cell_3.SetCellValue(kvp.Value.count_notsign);

                ICell cell_4 = row.CreateCell(3);
                cell_4.SetCellValue(kvp.Value.count_allowance);
                index += 1;
            }
        }

        private void 设置ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("功能未开放！");
        }

        

        public void SetOffday(List<bool> list)
        {
            for(int day = 0;day<list.Count;day++)
            {
                bool offday = list[day];
                Console.WriteLine(offday);
                MonthInfo.month_info[day].is_offday = list[day];
            }
            this.form_setoffday = null;

            LoadExcelData();
        }
    }
}
