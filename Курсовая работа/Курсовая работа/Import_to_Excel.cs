using Excel = Microsoft.Office.Interop.Excel;

namespace Курсовая_работа
{
    class Import_to_Excel : Results
    {
        public Import_to_Excel(string transit_time, double[] answersSecurity, double[] answersConsume, double[] answersCompetence, string answers_u, string full_name, int age, string sex, int number) : base(transit_time, answersSecurity, answersConsume, answersCompetence, answers_u, full_name, age, sex, number) { }

        public void data_excel()
        {
            Excel.Application ObjWorkExcel = (Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
            Excel.Workbook ObjWorkBook = ObjWorkExcel.ActiveWorkbook;
            Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1];

            for (int i = 6; i < 11; i++)
            {
                ObjWorkSheet.Cells[Test_number + 1, i].Value = AnswersSecurity[i - 6];
            }
            for (int i = 11; i < 15; i++)
            {
                ObjWorkSheet.Cells[Test_number + 1, i].Value = AnswersConsume[i - 11];
            }
            for (int i = 15; i < 20; i++)
            {
                ObjWorkSheet.Cells[Test_number + 1, i].Value = AnswersCompetence[i - 15];
            }
            for (int i = 20; i < 21; i++)
            {
                ObjWorkSheet.Cells[Test_number + 1, i].Value = AnswersConsume[i-16];
            }
            

            ObjWorkSheet.Cells[Test_number + 1, 21].FormulaLocal = "=(СУММ(F" + (Test_number + 1) + ":J" + (Test_number + 1) + "))/3,33*100";
            ObjWorkSheet.Cells[Test_number + 1, 22].FormulaLocal = "=(СУММ(K" + (Test_number + 1) + ":N" + (Test_number + 1) + ")+T" + (Test_number + 1) + ")/3,33*100";
            ObjWorkSheet.Cells[Test_number + 1, 23].FormulaLocal = "=(СУММ(O" + (Test_number + 1) + ":S" + (Test_number + 1) + "))/3,33*100";
            ObjWorkSheet.Cells[Test_number + 1, 24].Value = ItogIndex();
            ObjWorkSheet.Cells[Test_number + 1, 25].Value = "00:" + Transit_Time;
            ObjWorkSheet.Cells[Test_number + 1, 26].Value = level();

            ObjWorkBook.Save();
        }
    }
}
