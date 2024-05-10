using System;
using System.Drawing;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace Курсовая_работа
{
    public partial class SecondForm : Form
    {
        public SecondForm()
        {
            InitializeComponent();
            Excel.Application ObjWorkExcel = (Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
            Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkExcel.Sheets[1];

            int n = 0;
            while (ObjWorkSheet.Cells[n + 1, 1].Value != null)
            {
                n++;
            }
            dataGridView1.ColumnCount = 4;
            dataGridView1.RowCount = n;
            dataGridView1.Rows[0].DefaultCellStyle.BackColor = Color.AliceBlue;

            for (int i = 0; i < n; i++)
            {
                for (int j = 1; j <= 25; j++)
                {
                    if (j == 1)
                    {
                        dataGridView1.Rows[i].Cells[0].Value = ObjWorkSheet.Cells[i + 1, j].Text;
                    }
                    if (j == 2)
                    {
                        dataGridView1.Rows[i].Cells[1].Value = ObjWorkSheet.Cells[i + 1, j].Text;
                    }
                    if (j == 5)
                    {
                        dataGridView1.Rows[i].Cells[2].Value = ObjWorkSheet.Cells[i + 1, j].Value.ToString();
                    }
                    if (j == 24)
                    {
                        dataGridView1.Rows[i].Cells[3].Value = ObjWorkSheet.Cells[i + 1, j].Text;
                    }
                }
            }
            dataGridView1.Columns[0].Width = 55;
            dataGridView1.Columns[3].Width = 70;

            double sum = 0;
            for (int i = 1; i < n; i++)
            {
                sum += Convert.ToDouble(dataGridView1.Rows[i].Cells[3].Value);
                label1.Text = "Общий ИГЦН составляет " + Math.Round(sum / n, 2) + " балла(ов) из 10-ти";
            }

        }

        private void reveal_Click(object sender, EventArgs e)
        {
            Excel.Application ObjWorkExcel = (Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
            Excel.Workbook ObjWorkBook = ObjWorkExcel.ActiveWorkbook;
            Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1];

            int n = 0;
            while (ObjWorkSheet.Cells[n + 2, 24].Value != null)
            {
                n++;
            }

            if (n % 10 == 0)
            {
                for (int i = n - 8; i < n + 2; i++)
                {
                    chart1.Series[0].Points.AddXY(i - 1, ObjWorkSheet.Cells[i, 24].Value);
                    chart1.Series[0].Points[i - 2].Label = ObjWorkSheet.Cells[i, 24].Value.ToString();
                }
            }
            else if (n < 10)
            {
                for (int i = 2; i < n + 2; i++)
                {
                    chart1.Series[0].Points.AddXY(i - 1, ObjWorkSheet.Cells[i, 24].Value);
                    chart1.Series[0].Points[i - 2].Label = ObjWorkSheet.Cells[i, 24].Value.ToString();
                }
            }
            chart1.Visible = true;
        }
    }
}
