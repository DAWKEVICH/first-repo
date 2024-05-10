using System;
using Excel = Microsoft.Office.Interop.Excel;

namespace Курсовая_работа
{
    class Member : Testing
    {
        public string Full_name { get; set; }
        public int Age { get; set; }
        public string Sex { get; set; }

        public Member(string full_name, int age, string sex, int number) : base(number) { Full_name = full_name; Age = age; Sex = sex; }

        public void save()
        {
            Excel.Application ObjWorkExcel = (Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
            Excel.Workbook ObjWorkBook = ObjWorkExcel.ActiveWorkbook;
            Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1];

            ObjWorkSheet.Cells[Test_number + 1, 1].Value = Test_number;
            ObjWorkSheet.Cells[Test_number + 1, 2].Value = Full_name;
            ObjWorkSheet.Cells[Test_number + 1, 3].Value = Age;
            ObjWorkSheet.Cells[Test_number + 1, 4].Value = Sex;
            ObjWorkSheet.Cells[Test_number + 1, 5].Value = DateTime.Now;

            ObjWorkBook.Save();
        }
    }
}
