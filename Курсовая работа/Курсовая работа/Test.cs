using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Threading;

namespace Курсовая_работа
{
    public partial class Test : Form
    {
        List<Results> results = new List<Results>();
        List<Import_to_Excel> file_excel = new List<Import_to_Excel>();
        List<Import_to_Word> file_word = new List<Import_to_Word>();

        int n = 1;
        int t = 0;
        int m = 0, s = 0;

        public Test()
        {
            InitializeComponent();
            timer2.Start();
            chart1.Visible = false;
        }

        private void Further_Click(object sender, EventArgs e)
        {
            n++;
            Complete.Visible = false;

            if (n == 2)
            {
                if (answer1_1.Checked || answer1_2.Checked || answer1_3.Checked || answer1_4.Checked)
                    pb1.Visible = false;
                else
                    pb1.Visible = true;

                true2();
                false1();

                Back.Visible = true;
            }
            else if (n == 3)
            {
                if (answer2_1.Checked || answer2_2.Checked || answer2_3.Checked || answer2_4.Checked)
                    pb2.Visible = false;
                else
                    pb2.Visible = true;

                true3();
                false2();
            }
            else if (n == 4)
            {
                if (answer3_1.Checked || answer3_2.Checked || answer3_3.Checked || answer3_4.Checked)
                    pb3.Visible = false;
                else
                    pb3.Visible = true;

                true4();
                false3();
            }
            else if (n == 5)
            {
                if (answer4_1.Checked || answer4_2.Checked || answer4_3.Checked || answer4_4.Checked || answer4_5.Checked)
                    pb4.Visible = false;
                else
                    pb4.Visible = true;

                true5();
                false4();
            }
            else if (n == 6)
            {
                if (answer5_1.Checked || answer5_2.Checked || answer5_3.Checked)
                    pb5.Visible = false;
                else
                    pb5.Visible = true;

                true6();
                false5();
            }
            else if (n == 7)
            {
                if (answer6_1.Checked || answer6_2.Checked)
                    pb6.Visible = false;
                else
                    pb6.Visible = true;

                true7();
                false6();
            }
            else if (n == 8)
            {
                if (answer7_1.Checked || answer7_2.Checked || answer7_3.Checked)
                    pb7.Visible = false;
                else
                    pb7.Visible = true;

                true8();
                false7();
            }
            else if (n == 9)
            {
                if (answer8.Text != "")
                    pb8.Visible = false;
                else
                    pb8.Visible = true;

                true9();
                false8();
            }
            else if (n == 10)
            {
                if ((answer9_1.Text != "" && answer9_2.Text != "" && answer9_3.Text != "" && answer9_1.Text != answer9_2.Text && answer9_1.Text != answer9_3.Text && answer9_2.Text != answer9_3.Text) ||
                   (answer9_1.Text != "" && answer9_2.Text != "" && answer9_3.Text == "" && answer9_1.Text != answer9_2.Text) ||
                   (answer9_1.Text != "" && answer9_3.Text != "" && answer9_2.Text == "" && answer9_1.Text != answer9_3.Text) ||
                   (answer9_2.Text != "" && answer9_3.Text != "" && answer9_1.Text == "" && answer9_2.Text != answer9_3.Text))
                {
                    pb9.Visible = false;
                    true10();
                    false9();
                }
                else if ((answer9_1.Text == "" && answer9_2.Text == "") || (answer9_1.Text == "" && answer9_3.Text == "") || (answer9_2.Text == "" && answer9_3.Text == ""))
                {
                    pb9.Visible = true;
                    true10();
                    false9();
                }
                else
                {
                    n = 9;
                    mistake.Visible = true;
                    timer1.Start();
                }
            }
            else if (n == 11)
            {
                if (answer10_1.Checked || answer10_2.Checked || answer10_3.Checked)
                    pb10.Visible = false;
                else
                    pb10.Visible = true;

                true11();
                false10();
            }
            else if (n == 12)
            {
                if (answer11_1.Checked || answer11_2.Checked)
                    pb21.Visible = false;
                else
                    pb21.Visible = true;

                true12();
                false11();
            }
            else if (n == 13)
            {
                if (answer12_1.Checked || answer12_2.Checked || answer12_3.Checked)
                    pb22.Visible = false;
                else
                    pb22.Visible = true;

                true13();
                false12();
            }
            else if (n == 14)
            {
                if (answer13_1.Checked || answer13_2.Checked || answer13_3.Checked || answer13_4.Checked)
                    pb23.Visible = false;
                else
                    pb23.Visible = true;

                true14();
                false13();
            }
            else if (n == 15)
            {
                if (answer14_1.Checked || answer14_2.Checked || answer14_3.Checked || answer14_4.Checked)
                    pb24.Visible = false;
                else
                    pb24.Visible = true;

                true15();
                false14();

                Further.Visible = false;

                if (pb1.Visible == false && pb2.Visible == false && pb3.Visible == false && pb4.Visible == false && pb5.Visible == false && pb6.Visible == false && pb7.Visible == false && pb8.Visible == false && pb9.Visible == false && pb10.Visible == false && pb21.Visible == false && pb22.Visible == false && pb23.Visible == false && pb24.Visible == false)
                {
                    Complete.Visible = true;
                }
            }
        }

        private void Back_Click(object sender, EventArgs e)
        {
            n--;
            Complete.Visible = false;

            if (n == 1)
            {
                if (answer2_1.Checked || answer2_2.Checked || answer2_3.Checked || answer2_4.Checked)
                    pb2.Visible = false;
                else
                    pb2.Visible = true;

                true1();
                false2();

                Back.Visible = false;
            }
            else if (n == 2)
            {
                if (answer3_1.Checked || answer3_2.Checked || answer3_3.Checked || answer3_4.Checked)
                    pb3.Visible = false;
                else
                    pb3.Visible = true;

                true2();
                false3();
            }
            else if (n == 3)
            {
                if (answer4_1.Checked || answer4_2.Checked || answer4_3.Checked || answer4_4.Checked || answer4_5.Checked)
                    pb4.Visible = false;
                else
                    pb4.Visible = true;

                true3();
                false4();
            }
            else if (n == 4)
            {
                if (answer5_1.Checked || answer5_2.Checked || answer5_3.Checked)
                    pb5.Visible = false;
                else
                    pb5.Visible = true;

                true4();
                false5();
            }
            else if (n == 5)
            {
                if (answer6_1.Checked || answer6_2.Checked)
                    pb6.Visible = false;
                else
                    pb6.Visible = true;

                true5();
                false6();
            }
            else if (n == 6)
            {
                if (answer7_1.Checked || answer7_2.Checked || answer7_3.Checked)
                    pb7.Visible = false;
                else
                    pb7.Visible = true;

                true6();
                false7();
            }
            else if (n == 7)
            {
                if (answer8.Text != "")
                    pb8.Visible = false;
                else
                    pb8.Visible = true;

                true7();
                false8();
            }
            else if (n == 8)
            {
                if ((answer9_1.Text != "" && answer9_2.Text != "" && answer9_3.Text != "" && answer9_1.Text != answer9_2.Text && answer9_1.Text != answer9_3.Text && answer9_2.Text != answer9_3.Text) ||
                   (answer9_1.Text != "" && answer9_2.Text != "" && answer9_3.Text == "" && answer9_1.Text != answer9_2.Text) ||
                   (answer9_1.Text != "" && answer9_3.Text != "" && answer9_2.Text == "" && answer9_1.Text != answer9_3.Text) ||
                   (answer9_2.Text != "" && answer9_3.Text != "" && answer9_1.Text == "" && answer9_2.Text != answer9_3.Text))
                {
                    pb7.Visible = false;
                    true8();
                    false9();
                }
                else if ((answer9_1.Text == "" && answer9_2.Text == "") || (answer9_1.Text == "" && answer9_3.Text == "") || (answer9_2.Text == "" && answer9_3.Text == ""))
                {
                    pb9.Visible = true;
                    true8();
                    false9();
                }
                else
                {
                    n = 9;
                    mistake.Visible = true;
                    timer1.Start();
                }
            }
            else if (n == 9)
            {
                if (answer10_1.Checked || answer10_2.Checked || answer10_3.Checked)
                    pb10.Visible = false;
                else
                    pb10.Visible = true;

                true9();
                false10();
            }
            else if (n == 10)
            {
                if (answer11_1.Checked || answer11_2.Checked)
                    pb21.Visible = false;
                else
                    pb21.Visible = true;

                true10();
                false11();
            }
            else if (n == 11)
            {
                if (answer12_1.Checked || answer12_2.Checked || answer12_3.Checked)
                    pb22.Visible = false;
                else
                    pb22.Visible = true;

                true11();
                false12();
            }
            else if (n == 12)
            {
                if (answer13_1.Checked || answer13_2.Checked || answer13_3.Checked || answer14_3.Checked)
                    pb23.Visible = false;
                else
                    pb23.Visible = true;

                true12();
                false13();
            }
            else if (n == 13)
            {
                if (answer14_1.Checked || answer14_2.Checked || answer14_3.Checked || answer14_4.Checked)
                    pb24.Visible = false;
                else
                    pb24.Visible = true;

                true13();
                false14();
            }
            else if (n == 14)
            {
                if (answer15_1.Checked || answer15_2.Checked || answer15_3.Checked || answer15_4.Checked)
                    pb25.Visible = false;
                else
                    pb25.Visible = true;

                true14();
                false15();

                Further.Visible = true;
            }
        }

        private void Clear_Click(object sender, EventArgs e)
        {
            if (n == 1)
            {
                answer1_1.Checked = false;
                answer1_2.Checked = false;
                answer1_3.Checked = false;
                answer1_4.Checked = false;

                pb1.Visible = true;
            }
            else if (n == 2)
            {
                answer2_1.Checked = false;
                answer2_2.Checked = false;
                answer2_3.Checked = false;
                answer2_4.Checked = false;

                pb2.Visible = true;
            }
            else if (n == 3)
            {
                answer3_1.Checked = false;
                answer3_2.Checked = false;
                answer3_3.Checked = false;
                answer3_4.Checked = false;

                pb3.Visible = true;
            }
            else if (n == 4)
            {
                answer4_1.Checked = false;
                answer4_2.Checked = false;
                answer4_3.Checked = false;
                answer4_4.Checked = false;
                answer4_5.Checked = false;

                pb4.Visible = true;
            }
            else if (n == 5)
            {
                answer5_1.Checked = false;
                answer5_2.Checked = false;
                answer5_3.Checked = false;

                pb5.Visible = true;
            }
            else if (n == 6)
            {
                answer6_1.Checked = false;
                answer6_2.Checked = false;

                pb6.Visible = true;
            }
            else if (n == 7)
            {
                answer7_1.Checked = false;
                answer7_2.Checked = false;
                answer7_3.Checked = false;

                pb7.Visible = true;
            }
            else if (n == 8)
            {
                answer8.Clear();

                pb8.Visible = true;
            }
            else if (n == 9)
            {
                answer9_1.Text = "";
                answer9_2.Text = "";
                answer9_3.Text = "";

                pb9.Visible = true;
            }
            else if (n == 10)
            {
                answer10_1.Checked = false;
                answer10_2.Checked = false;
                answer10_3.Checked = false;

                pb10.Visible = true;

            }
            else if (n == 11)
            {
                answer11_1.Checked = false;
                answer11_2.Checked = false;

                pb11.Visible = true;
            }
            else if (n == 12)
            {
                answer12_1.Checked = false;
                answer12_2.Checked = false;
                answer12_3.Checked = false;

                pb12.Visible = true;
            }
            else if (n == 13)
            {
                answer13_1.Checked = false;
                answer13_2.Checked = false;
                answer13_3.Checked = false;
                answer13_4.Checked = false;

                pb13.Visible = true;
            }
            else if (n == 14)
            {
                answer14_1.Checked = false;
                answer14_2.Checked = false;
                answer14_3.Checked = false;
                answer14_4.Checked = false;

                pb14.Visible = true;
            }
            else if (n == 15)
            {
                answer15_1.Checked = false;
                answer15_2.Checked = false;
                answer15_3.Checked = false;
                answer15_4.Checked = false;

                pb15.Visible = true;
                Complete.Visible = false;
            }
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            t++;
            if (t == 20)
            {
                mistake.Visible = false;
                t = 0;
                timer1.Stop();
            }
        }

        private void Complete_Click(object sender, EventArgs e)
        {
            pb25.Visible = true;
            Thread.Sleep(200);
            pb21.Visible = false;
            pb22.Visible = false;
            pb23.Visible = false;
            pb24.Visible = false;
            pb25.Visible = false;
            pb26.Visible = false;
            pb27.Visible = false;
            pb28.Visible = false;
            pb29.Visible = false;
            pb30.Visible = false;

            timer2.Stop();
            Excel.Application ObjWorkExcel = (Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
            Excel.Workbook ObjWorkBook = ObjWorkExcel.ActiveWorkbook;
            Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1];

            int n = 0;
            while (ObjWorkSheet.Cells[n + 2, 1].Value != null)
            {
                n++;
            }

            string full_name = ObjWorkSheet.Cells[n + 1, 2].Value;
            int age = Convert.ToInt32(ObjWorkSheet.Cells[n + 1, 3].Value);
            string sex = ObjWorkSheet.Cells[n + 1, 4].Value;


            /////////////////////////ANSWERS//////////////////////////
            double[] answersSecurity = new double[5];
            double[] answersConsume = new double[5];
            double[] answersCompetence = new double[5];
            string answers_u = "";

            //Question 1
            if (answer1_4.Checked == true)
            {
                answersSecurity[0] = 0.666;
                answers_u += "4;";
            }
            else
            {
                answersSecurity[0] = 0;
                if (answer1_1.Checked == true)
                {
                    answers_u += "1;";
                }
                else if (answer1_2.Checked == true)
                {
                    answers_u += "2;";
                }
                else answers_u += "3;";
            }

            //Question 2
            if (answer2_3.Checked == true)
            {
                answersSecurity[1] = 0.666;
                answers_u += "3;";
            }
            else
            {
                answersSecurity[1] = 0;
                if (answer2_1.Checked == true)
                {
                    answers_u += "1;";
                }
                else if (answer2_2.Checked == true)
                {
                    answers_u += "2;";
                }
                else answers_u += "3;";
            }

            //Question 3
            if (answer3_3.Checked == true)
            {
                answersSecurity[2] = 0.666;
                answers_u += "3;";
            }
            else
            {
                answersSecurity[2] = 0;
                if (answer3_1.Checked == true)
                {
                    answers_u += "1;";
                }
                else if (answer3_2.Checked == true)
                {
                    answers_u += "2;";
                }
                else answers_u += "3;";
            }

            //Question 4
            if (answer4_4.Checked == true)
            {
                answersSecurity[3] = 0.666;
                answers_u += "4;";
            }
            else
            {
                answersSecurity[3] = 0;
                if (answer4_1.Checked == true)
                {
                    answers_u += "1;";
                }
                else if (answer4_2.Checked == true)
                {
                    answers_u += "2;";
                }
                else if (answer4_3.Checked == true)
                {
                    answers_u += "3;";
                }
                else answers_u += "5;";
            }

            //Question 5
            if (answer5_2.Checked == true)
            {
                answersSecurity[4] = 0.666;
                answers_u += "2;";
            }
            else
            {
                answersSecurity[4] = 0;
                if (answer5_1.Checked == true)
                {
                    answers_u += "1;";
                }
                else answers_u += "3;";
            }

            //Question 6
            if (answer6_1.Checked == true)
            {
                answersConsume[0] = 0.666;
                answers_u += "1;";
            }
            else
            {
                answersConsume[0] = 0;
                answers_u += "1;";
            }

            //Question 7
            if (answer7_2.Checked == true)
            {
                answersConsume[1] = 0.666;
                answers_u += "2;";
            }
            else
            {
                answersConsume[1] = 0;
                if (answer7_1.Checked == true)
                {
                    answers_u += "1;";
                }
                else answers_u += "3;";
            }

            //Question 8
            answers_u += answer8.Text + ";";
            string answer = answer8.Text.ToLower();
            answer = answer.Replace(" ", string.Empty);
            if (answer == "интернетвещей")
            {
                answersConsume[2] = 0.666;
            }
            else
            {
                answersConsume[2] = 0;
            }

            //Question 9
            if (answer9_1.Text == "1" && answer9_2.Text == "2" && answer9_3.Text == "3")
            {
                answersConsume[3] = 0.666;
            }
            else if (answer9_1.Text == "2" || answer9_2.Text == "1" || answer9_3.Text == "3")
            {
                answersConsume[3] = 0.333;
            }
            else
            {
                answersConsume[3] = 0;
            }
            answers_u += answer9_1.Text + answer9_2.Text + answer9_3.Text + ";";

            //Question 10
            if (answer10_3.Checked == true)
            {
                answersCompetence[0] = 0.666;
                answers_u += "3;";
            }
            else
            {
                answersCompetence[0] = 0;
                if (answer10_1.Checked == true)
                {
                    answers_u += "1;";
                }
                else answers_u += "2;";
            }

            //Question 11
            if (answer11_1.Checked == true)
            {
                answersCompetence[1] = 0.666;
                answers_u += "1;";
            }
            else
            {
                answersCompetence[1] = 0;
                answers_u += "2;";
            }

            //Question 12
            if (answer12_3.Checked == true)
            {
                answersCompetence[2] = 0.666;
                answers_u += "3;";
            }
            else
            {
                answersCompetence[2] = 0;
                if (answer12_1.Checked == true)
                {
                    answers_u += "1;";
                }
                else answers_u += "2;";
            }

            //Question 13
            if (answer13_4.Checked == true)
            {
                answersCompetence[3] = 0.666;
                answers_u += "4;";
            }
            else
            {
                answersCompetence[3] = 0;
                if (answer13_1.Checked == true)
                {
                    answers_u += "1;";
                }
                else if (answer13_2.Checked == true)
                {
                    answers_u += "2;";
                }
                else answers_u += "3;";
            }

            //Question 14
            if (answer14_2.Checked == true)
            {
                answersCompetence[4] = 0.666;
                answers_u += "2;";
            }
            else
            {
                answersCompetence[4] = 0;
                if (answer14_1.Checked == true)
                {
                    answers_u += "1;";
                }
                else if (answer14_3.Checked == true)
                {
                    answers_u += "3;";
                }
                else answers_u += "4;";
            }

            //Question 15
            if (answer15_2.Checked == true)
            {
                answersConsume[4] = 0.666;
                answers_u += "2;";
            }
            else
            {
                answersConsume[4] = 0;
                if (answer15_1.Checked == true)
                {
                    answers_u += "1;";
                }
                else if (answer15_3.Checked == true)
                {
                    answers_u += "3;";
                }
                else answers_u += "4;";
            }

            ///////////////////////////////////////////////////////////
            results.Add(new Results(time.Text, answersSecurity, answersConsume, answersCompetence, answers_u, full_name, age, sex, n));
            results[results.Count - 1].processing_of_results(FIO, PointsSecurity, PointsConsume, PointsCompetence, Lvl, Date);
            false15();
            panel11.Visible = true;
            FIO.Visible = true;
            PointsSecurity.Visible = true;
            PointsConsume.Visible = true;
            PointsCompetence.Visible = true;
            Date.Visible = true;
            PrintCerteficate.Visible = true;
            chart1.Visible = true;

            chart1.Series.Clear();
            chart1.Titles.Add("Диаграмма результатов");
            chart1.Series.Add("Pie");
            chart1.Series["Pie"].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Pie;
            double[] yValues = { results[results.Count - 1].resultSecurityForPie(), results[results.Count - 1].resultConsumeForPie(), results[results.Count - 1].resultCompetenceForPie() };
            string[] xValues = { "Безопасность", "Потребление", "Компетенции" };
            chart1.Series["Pie"].Points.DataBindXY(xValues, yValues);
            chart1.Series["Pie"]["PieLabelStyle"] = "Disabled";

            file_excel.Add(new Import_to_Excel(time.Text, answersSecurity, answersConsume, answersCompetence, answers_u, full_name, age, sex, n));
            file_excel[file_excel.Count - 1].data_excel();

            file_word.Add(new Import_to_Word(time.Text, answersSecurity, answersConsume, answersCompetence, answers_u, full_name, age, sex, n));
            file_word[file_word.Count - 1].data_word();
        }

        private void PrintCerteficate_Click(object sender, EventArgs e)
        {
            file_word[file_word.Count - 1].data_word_template();
            MessageBox.Show("Сертификат успешно сохранен в корневой папке проекта по пути: Курсовая работа\\Курсовая работа\\bin\\Debug\\Сертификаты", "Сохранение", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        /////////////////////////////////////////////////////////////////////////////METHODS FOR OPENING AND HIDING QESTIONS//////////////////////////////////////////////////////////////////////////////////////
        public void true1()
        {
            panel1.Visible = true;
            question1.Visible = true;
            answer1_1.Visible = true;
            answer1_2.Visible = true;
            answer1_3.Visible = true;
            answer1_4.Visible = true;
        }

        public void true2()
        {
            panel2.Visible = true;
            question2.Visible = true;
            answer2_1.Visible = true;
            answer2_2.Visible = true;
            answer2_3.Visible = true;
            answer2_4.Visible = true;
        }

        public void true3()
        {
            panel3.Visible = true;
            question3.Visible = true;
            answer3_1.Visible = true;
            answer3_2.Visible = true;
            answer3_3.Visible = true;
            answer3_4.Visible = true;
        }

        public void true4()
        {
            panel4.Visible = true;
            question4.Visible = true;
            answer4_1.Visible = true;
            answer4_2.Visible = true;
            answer4_3.Visible = true;
            answer4_4.Visible = true;
            answer4_5.Visible = true;
        }

        public void true5()
        {
            panel5.Visible = true;
            question5.Visible = true;
            answer5_1.Visible = true;
            answer5_2.Visible = true;
            answer5_3.Visible = true;
        }

        public void true6()
        {
            panel6.Visible = true;
            question6.Visible = true;
            answer6_1.Visible = true;
            answer6_2.Visible = true;
        }

        public void true7()
        {
            panel7.Visible = true;
            question7.Visible = true;
            answer7_1.Visible = true;
            answer7_2.Visible = true;
            answer7_3.Visible = true;
        }

        public void true8()
        {
            panel8.Visible = true;
            question8.Visible = true;
            answer8.Visible = true;
        }

        public void true9()
        {
            panel9.Visible = true;
            question9.Visible = true;
            p1.Visible = true;
            p2.Visible = true;
            p3.Visible = true;
            answer9_1.Visible = true;
            answer9_2.Visible = true;
            answer9_3.Visible = true;
            o1.Visible = true;
            o2.Visible = true;
            o3.Visible = true;
        }

        public void true10()
        {
            panel10.Visible = true;
            question10.Visible = true;
            answer10_1.Visible = true;
            answer10_2.Visible = true;
            answer10_3.Visible = true;
        }
        public void true11()
        {
            panel112.Visible = true;
            question11.Visible = true;
            answer11_1.Visible = true;
            answer11_2.Visible = true;
        }
        public void true12()
        {
            panel12.Visible = true;
            question12.Visible = true;
            answer12_1.Visible = true;
            answer12_2.Visible = true;
            answer12_3.Visible = true;
        }
        public void true13()
        {
            panel13.Visible = true;
            question13.Visible = true;
            answer13_1.Visible = true;
            answer13_2.Visible = true;
            answer13_3.Visible = true;
            answer13_4.Visible = true;
        }
        public void true14()
        {
            panel14.Visible = true;
            question14.Visible = true;
            answer14_1.Visible = true;
            answer14_2.Visible = true;
            answer14_3.Visible = true;
            answer14_4.Visible = true;
        }
        public void true15()
        {
            panel15.Visible = true;
            question15.Visible = true;
            answer15_1.Visible = true;
            answer15_2.Visible = true;
            answer15_3.Visible = true;
            answer15_4.Visible = true;
        }

        public void false1()
        {
            panel1.Visible = false;
            question1.Visible = false;
            answer1_1.Visible = false;
            answer1_2.Visible = false;
            answer1_3.Visible = false;
            answer1_4.Visible = false;
        }

        public void false2()
        {
            panel2.Visible = false;
            question2.Visible = false;
            answer2_1.Visible = false;
            answer2_2.Visible = false;
            answer2_3.Visible = false;
            answer2_4.Visible = false;
        }

        public void false3()
        {
            panel3.Visible = false;
            question3.Visible = false;
            answer3_1.Visible = false;
            answer3_2.Visible = false;
            answer3_3.Visible = false;
            answer3_4.Visible = false;
        }

        public void false4()
        {
            panel4.Visible = false;
            question4.Visible = false;
            answer4_1.Visible = false;
            answer4_2.Visible = false;
            answer4_3.Visible = false;
            answer4_4.Visible = false;
            answer4_5.Visible = false;
        }

        public void false5()
        {
            panel5.Visible = false;
            question5.Visible = false;
            answer5_1.Visible = false;
            answer5_2.Visible = false;
            answer5_3.Visible = false;
        }

        public void false6()
        {
            panel6.Visible = false;
            question6.Visible = false;
            answer6_1.Visible = false;
            answer6_2.Visible = false;
        }

        public void false7()
        {
            panel7.Visible = false;
            question7.Visible = false;
            answer7_1.Visible = false;
            answer7_2.Visible = false;
            answer7_3.Visible = false;
        }

        public void false8()
        {
            panel8.Visible = false;
            question8.Visible = false;
            answer8.Visible = false;
        }

        public void false9()
        {
            panel9.Visible = false;
            question9.Visible = false;
            p1.Visible = false;
            p2.Visible = false;
            p3.Visible = false;
            answer9_1.Visible = false;
            answer9_2.Visible = false;
            answer9_3.Visible = false;
            o1.Visible = false;
            o2.Visible = false;
            o3.Visible = false;
        }

        public void false10()
        {
            panel10.Visible = false;
            question10.Visible = false;
            answer10_1.Visible = false;
            answer10_2.Visible = false;
            answer10_3.Visible = false;
        }

        public void false11()
        {
            panel112.Visible = false;
            question11.Visible = false;
            answer11_1.Visible = false;
            answer11_2.Visible = false;
        }
        public void false12()
        {
            panel12.Visible = false;
            question12.Visible = false;
            answer12_1.Visible = false;
            answer12_2.Visible = false;
            answer12_3.Visible = false;
        }
        public void false13()
        {
            panel13.Visible = false;
            question13.Visible = false;
            answer13_1.Visible = false;
            answer13_2.Visible = false;
            answer13_3.Visible = false;
            answer13_4.Visible = false;
        }
        public void false14()
        {
            panel14.Visible = false;
            question14.Visible = false;
            answer14_1.Visible = false;
            answer14_2.Visible = false;
            answer14_3.Visible = false;
            answer14_4.Visible = false;
        }
        public void false15()
        {
            panel15.Visible = false;
            question15.Visible = false;
            answer15_1.Visible = false;
            answer15_2.Visible = false;
            answer15_3.Visible = false;
            answer15_4.Visible = false;
        }

        private void answer8_KeyPress_1(object sender, KeyPressEventArgs e)
        {
            if (!Char.IsLetter(e.KeyChar) && e.KeyChar != 8 && e.KeyChar != 32)
            {
                e.Handled = true;
            }
        }

        private void answer15_1_Click(object sender, EventArgs e)
        {
            if (pb1.Visible == false && pb2.Visible == false && pb3.Visible == false && pb4.Visible == false && pb5.Visible == false && pb6.Visible == false && pb7.Visible == false && pb8.Visible == false && pb9.Visible == false && pb10.Visible == false && pb21.Visible == false && pb22.Visible == false && pb23.Visible == false && pb24.Visible == false && pb25.Visible == false)
            {
                Complete.Visible = true;
            }
        }

        private void answer15_2_Click(object sender, EventArgs e)
        {
            if (pb1.Visible == false && pb2.Visible == false && pb3.Visible == false && pb4.Visible == false && pb5.Visible == false && pb6.Visible == false && pb7.Visible == false && pb8.Visible == false && pb9.Visible == false && pb10.Visible == false && pb21.Visible == false && pb22.Visible == false && pb23.Visible == false && pb24.Visible == false && pb25.Visible == false)
            {
                Complete.Visible = true;
            }
        }

        private void answer15_3_Click(object sender, EventArgs e)
        {
            if (pb1.Visible == false && pb2.Visible == false && pb3.Visible == false && pb4.Visible == false && pb5.Visible == false && pb6.Visible == false && pb7.Visible == false && pb8.Visible == false && pb9.Visible == false && pb10.Visible == false && pb21.Visible == false && pb22.Visible == false && pb23.Visible == false && pb24.Visible == false && pb25.Visible == false)
            {
                Complete.Visible = true;
            }
        }

        private void answer15_4_Click(object sender, EventArgs e)
        {
            if (pb1.Visible == false && pb2.Visible == false && pb3.Visible == false && pb4.Visible == false && pb5.Visible == false && pb6.Visible == false && pb7.Visible == false && pb8.Visible == false && pb9.Visible == false && pb10.Visible == false && pb21.Visible == false && pb22.Visible == false && pb23.Visible == false && pb24.Visible == false && pb25.Visible == false)
            {
                Complete.Visible = true;
            }
        }

        private void PrintResults_Click(object sender, EventArgs e)
        {
            file_word[file_word.Count - 1].print_the_data();
            MessageBox.Show("Результат успешно сохранен в корневой папке проекта по пути: Курсовая работа\\Курсовая работа\\bin\\Debug\\Результаты", "Сохранение", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }


        private void timer2_Tick(object sender, EventArgs e)
        {
            s = s + 1;

            if (s == 60)
            {
                m = m + 1;
                s = 0;
            }

            if (s < 10)
            {
                if (m < 10)
                {
                    time.Text = "0" + Convert.ToString(m) + ":0" + Convert.ToString(s);
                }
                else
                {
                    time.Text = Convert.ToString(m) + ":0" + Convert.ToString(s);
                }
            }
            else if (s > 10)
            {
                if (m < 10)
                {
                    time.Text = "0" + Convert.ToString(m) + ":" + Convert.ToString(s);
                }
                else
                {
                    time.Text = Convert.ToString(m) + ":" + Convert.ToString(s);
                }
            }
        }
    }
}
