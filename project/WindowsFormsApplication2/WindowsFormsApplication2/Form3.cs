using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApplication2
{
    public partial class Form3 : Form
    {

        List<bool> offday = new List<bool>();
        List<CheckBox> all_checkbox = new List<CheckBox>();
        Form1 main_form;
        public Form3(Form1 formmain)
        {
            main_form = formmain;
            InitializeComponent();
        }

        public void ShowCheckBoxs(DateTime date_time, List<bool> offday_list)
        {
            int day_count = DateTime.DaysInMonth(date_time.Year, date_time.Month);
            for (int i = 0; i < day_count; i++)
            {
                CheckBox checkbox = new CheckBox();

                checkbox.AutoSize = true;
                checkbox.Location = new System.Drawing.Point(7+120*(i/10), 25*(1+i%10));
                checkbox.Name = "checkBox"+(i+1).ToString();
                checkbox.Size = new System.Drawing.Size(78, 16);
                checkbox.TabIndex = 0;
                checkbox.Text = "checkBox" + (i + 1).ToString();
                checkbox.UseVisualStyleBackColor = true;
                checkbox.Parent = this.groupBox1;
                checkbox.Checked = offday_list[i];
                DateTime time = new DateTime(date_time.Year, date_time.Month, i + 1);
                string str = "";
                if (time.DayOfWeek == DayOfWeek.Sunday || time.DayOfWeek == DayOfWeek.Saturday)
                {
                    //checkbox.BackColor = Color.Blue;
                    checkbox.ForeColor = Color.Red;
                    str = "周末";
                }
                checkbox.Text = string.Format("{0}-{1}{2}", date_time.Month, i + 1,str);
                checkbox.Show();
                checkbox.CheckedChanged += new System.EventHandler(this.checkBox_CheckedChanged);

                all_checkbox.Add(checkbox);
            }
            
        }


        private void progressBar1_Click(object sender, EventArgs e)
        {
       
        }

        private void checkBox_CheckedChanged(object sender, EventArgs e)
        {
            offday.Clear();
            for(int i = 0;i< all_checkbox.Count;i++)
            {
                offday.Add(all_checkbox[i].Checked);
            }
        }

        public void OnClose(object sender, FormClosedEventArgs e)
        {
            main_form.SetOffday(offday);
        }

        private void Form3_Load(object sender, EventArgs e)
        {

        }
    }
}
