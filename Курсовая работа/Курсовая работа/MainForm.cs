using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Diagnostics;
using Excel = Microsoft.Office.Interop.Excel;

namespace Курсовая_работа
{
    public partial class MainForm : Form
    {
        List<Testing> dictation = new List<Testing>();
        List<Member> participant = new List<Member>();

        bool flag = false;

        public int Number { get; set; }

        public MainForm()
        {
            InitializeComponent();
            Process.Start(System.IO.Path.GetFullPath(@"БД опрошенных по тесту.xlsx"));
            Thread.Sleep(500);
            наГлавнуюToolStripMenuItem.BackColor = Color.FromArgb(175, 202, 228);
            date.Text = DateTime.Now.ToString("dd.MM.yyyy");
        }

        private void наГлавнуюToolStripMenuItem_Click(object sender, EventArgs e)
        {
            помощьToolStripMenuItem.BackColor = Color.Transparent;
            тестированиеToolStripMenuItem.BackColor = Color.Transparent;
            результатыОпросовToolStripMenuItem.BackColor = Color.Transparent;
            наГлавнуюToolStripMenuItem.BackColor = Color.FromArgb(175, 202, 228);
            тестированиеToolStripMenuItem.Enabled = false;
            помощьToolStripMenuItem.Enabled = true;
            результатыОпросовToolStripMenuItem.Enabled = true;
            checking();
            int i = 0;
            if (MdiChildren.Count() > 0)
            {
                do
                {
                    MdiChildren[i].Close();
                }
                while (MdiChildren.Count() > 0);
            }
            flag = false;
            visible_true();
        }

        private void помощьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            наГлавнуюToolStripMenuItem.BackColor = Color.Transparent;
            результатыОпросовToolStripMenuItem.BackColor = Color.Transparent;
            помощьToolStripMenuItem.BackColor = Color.FromArgb(175, 202, 228);
            visible_false();
            pictureBox1.Visible = true;
            pictureBox3.Visible = true;
            label7.Visible = true;
            button1.Visible = true;
        }

        private void результатыОпросовToolStripMenuItem_Click(object sender, EventArgs e)
        {
            наГлавнуюToolStripMenuItem.BackColor = Color.Transparent;
            помощьToolStripMenuItem.BackColor = Color.Transparent;
            результатыОпросовToolStripMenuItem.BackColor = Color.FromArgb(175, 202, 228);
            if (this.MdiChildren.Any())
            {
                return;
            }
            else
            {
                visible_false();
                SecondForm MdiChild = new SecondForm();
                MdiChild.MdiParent = this;
                MdiChild.Dock = DockStyle.Fill;
                MdiChild.Show();
            }
        }

        private void выходToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            Excel.Application ObjWorkExcel = (Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
            ObjWorkExcel.Visible = false;
        }

        private void full_name_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!Char.IsLetter(e.KeyChar) && e.KeyChar != 8 && e.KeyChar != 32)
            {
                e.Handled = true;
            }
        }

        private void full_name_TextChanged(object sender, EventArgs e)
        {
            if (full_name.Text.Length > 3 && age.Value > 6 && (male.Checked == true || female.Checked == true))
            {
                start_test.Enabled = true;
            }

            if (flag == false)
            {
                search();
                flag = true;
            }
        }

        private void age_ValueChanged(object sender, EventArgs e)
        {
            if (full_name.Text.Length > 3 && age.Value > 6 && (male.Checked == true || female.Checked == true))
            {
                start_test.Enabled = true;
            }
        }

        private void male_CheckedChanged(object sender, EventArgs e)
        {
            if (full_name.Text.Length > 3 && age.Value > 6 && (male.Checked == true || female.Checked == true))
            {
                start_test.Enabled = true;
            }
        }

        private void female_CheckedChanged(object sender, EventArgs e)
        {
            if (full_name.Text.Length > 3 && age.Value > 6 && (male.Checked == true || female.Checked == true))
            {
                start_test.Enabled = true;
            }
        }

        private void start_test_Click(object sender, EventArgs e)
        {
            string sex = "";
            if (male.Checked)
            {
                sex = "муж.";
            }
            else
            {
                sex = "жен.";
            }
            participant.Add(new Member(full_name.Text, Convert.ToInt32(age.Value), sex, Number));
            participant[participant.Count - 1].save();

            full_name.Clear();
            age.Value = 18;
            female.Checked = false;
            male.Checked = true;
            наГлавнуюToolStripMenuItem.BackColor = Color.Transparent;
            тестированиеToolStripMenuItem.BackColor = Color.FromArgb(175, 202, 228);
            помощьToolStripMenuItem.Enabled = false;
            результатыОпросовToolStripMenuItem.Enabled = false;
            тестированиеToolStripMenuItem.Enabled = true;
            visible_false();
            Test MdiChild = new Test();
            MdiChild.MdiParent = this;
            MdiChild.Dock = DockStyle.Fill;
            MdiChild.Show();
        }

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            try
            {
                Excel.Application ObjWorkExcel = (Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
                checking();
                ObjWorkExcel.Quit();
            }
            catch
            {
                this.Close();
                CloseProcess(); // ПРОВЕРЕНО
            }
        }

        // ПРОВЕРЕНО // Завершаем все процессы Экселя и Ворда, чтобы они не оставались в памяти после закрытия приложения
        public void CloseProcess()
        {
            Process[] ListExcel;
            Process[] ListWord;
            ListExcel = Process.GetProcessesByName("EXCEL");
            ListWord = Process.GetProcessesByName("WINWORD");
            foreach (Process proc in ListExcel)
            {
                proc.Kill();
            }
            foreach (Process proc in ListWord)
            {
                proc.Kill();
            }
        }

        public void search_number()
        {
            Excel.Application ObjWorkExcel = (Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
            Excel.Workbook ObjWorkBook = ObjWorkExcel.ActiveWorkbook;
            Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1];

            int n = 0;
            while (ObjWorkSheet.Cells[n + 2, 1].Value != null)
            {
                n++;
            }
            dictation.Add(new Testing(n + 1));
            Number = n + 1;
        }

        public async void search()
        {
            await Task.Run(() => search_number());
        }

        private void checking()
        {
            Excel.Application ObjWorkExcel = (Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
            Excel.Workbook ObjWorkBook = ObjWorkExcel.ActiveWorkbook;
            Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1];

            int n = 0;
            while (ObjWorkSheet.Cells[n + 1, 1].Value != null)
            {
                n++;
            }

            if (ObjWorkExcel.Cells[n, 1].Text != "" && ObjWorkExcel.Cells[n, 6].Text == "")
            {
                for (int i = 1; i < 6; i++)
                {
                    ObjWorkSheet.Cells[n, i].Value = "";
                }
            }

            ObjWorkBook.Save();
        }

        public void visible_true()
        {
            pictureBox3.Visible = false;
            label7.Visible = false;
            button1.Visible = false;
            pictureBox1.Visible = true;
            label1.Visible = true;
            label2.Visible = true;
            label3.Visible = true;
            label4.Visible = true;
            label5.Visible = true;
            label6.Visible = true;
            date.Visible = true;
            full_name.Visible = true;
            age.Visible = true;
            male.Visible = true;
            female.Visible = true;
            start_test.Visible = true;
            pictureBox2.Visible = true;
        }

        public void visible_false()
        {
            label1.Visible = false;
            label2.Visible = false;
            label3.Visible = false;
            label4.Visible = false;
            label5.Visible = false;
            label6.Visible = false;
            date.Visible = false;
            full_name.Visible = false;
            age.Visible = false;
            male.Visible = false;
            female.Visible = false;
            start_test.Visible = false;
            pictureBox2.Visible = false;
            pictureBox3.Visible = false;
            button1.Visible = false;
            label7.Visible = false;
            pictureBox1.Visible = false;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Clipboard.SetText("готовкцифре.рф/navigator", TextDataFormat.UnicodeText);
        }
    }
}
