using System;
using System.Windows.Forms;

namespace Курсовая_работа
{
    class Results : Member
    {
        public string Transit_Time { get; set; }
        public double[] AnswersSecurity { get; set; }
        public double[] AnswersConsume { get; set; }
        public double[] AnswersCompetence { get; set; }
        public string Answers_u { get; set; }

        public Results(string transit_time, double[] answersSecurity, double[] answersConsume, double[] answersCompetence, string answers_u, string full_name, int age, string sex, int number) : base(full_name, age, sex, number) { Transit_Time = transit_time; AnswersSecurity = answersSecurity; AnswersConsume = answersConsume; AnswersCompetence = answersCompetence; Answers_u = answers_u; }

        public void processing_of_results(Label FIO, Label PointsSecurity, Label PointsConsume, Label PointsCompetence, Label Lvl, Label Date)
        {
            FIO.Text = Full_name;
            PointsSecurity.Text = resultSecurityForPie().ToString() + " б.";
            PointsConsume.Text = resultConsumeForPie().ToString() + " б.";
            PointsCompetence.Text = resultCompetenceForPie().ToString() + " б.";
            Date.Text = DateTime.Now.ToString("dd.MM.yyyy") + " г.";
            Lvl.Text = level() + ", Ваш ИЦГН - " + Math.Round(ItogIndex(), 2).ToString() + " балла(ов)";
        }

        public double resultSecurity()
        {
            double sum1 = 0;
            for (int i = 0; i < 5; i++)
            {
                sum1 += AnswersSecurity[i];
            }
            double itogSecurity = sum1 / 3.33;
            return Math.Round(itogSecurity * 100, 2);
        }

        public double resultSecurityForPie()
        {
            double sum1 = 0;
            for (int i = 0; i < 5; i++)
            {
                sum1 += AnswersSecurity[i];
            }
            return Math.Round(sum1, 2);
        }
        public double resultConsumeForPie()
        {
            double sum2 = 0;
            for (int i = 0; i < 5; i++)
            {
                sum2 += AnswersConsume[i];
            }
            return Math.Round(sum2, 2);
        }

        public double resultCompetenceForPie()
        {
            double sum3 = 0;
            for (int i = 0; i < 5; i++)
            {
                sum3 += AnswersCompetence[i];
            }
            return Math.Round(sum3, 2);
        }
        public double resultConsume()
        {
            double sum2 = 0;
            for (int i = 0; i < 5; i++)
            {
                sum2 += AnswersConsume[i];
            }
            double itogConsume = sum2 / 3.33;
            return Math.Round(itogConsume * 100, 2);
        }
        public double resultCompetence()
        {
            double sum3 = 0;
            for (int i = 0; i < 5; i++)
            {
                sum3 += AnswersCompetence[i];
            }
            double itogCompetence = sum3 / 3.33;
            return Math.Round(itogCompetence * 100, 2);
        }

        public double ItogIndex()
        {
            double itogo = (resultSecurity() + resultConsume() + resultCompetence()) / 30.0;
            if (itogo == 9.99)
            {
                return 10;
            }
            else
                return Math.Round(itogo, 2);
        }

        public string level()
        {

            string[] transit_time = Transit_Time.Split(':');
            string lvl = "";
            if (Convert.ToInt32(transit_time[0]) < 5 && ItogIndex() >= 9 && resultSecurity() >= 85 && resultConsume() >= 85 && resultCompetence() >= 85)
            {
                lvl = "продвинутый";
            }
            else if (Convert.ToInt32(transit_time[0]) <= 10 && ItogIndex() >= 7 && ItogIndex() <= 100 && resultSecurity() > 65 && resultConsume() > 65 && resultCompetence() > 65)
            {
                lvl = "уверенный";
            }
            else
            {
                lvl = "начинающий";
            }
            return lvl;
        }
    }
}
