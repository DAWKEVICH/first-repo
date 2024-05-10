using System;
using System.Diagnostics;
using Microsoft.Office.Interop.Word;
using Application = Microsoft.Office.Interop.Word.Application;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace Курсовая_работа
{
    class Import_to_Word : Results
    {
        public Import_to_Word(string transit_time, double[] answersSecurity, double[] answersConsume, double[] answersCompetence, string answers_u, string full_name, int age, string sex, int number) : base(transit_time, answersSecurity, answersConsume, answersCompetence, answers_u, full_name, age, sex, number) { }

        public void data_word()
        {
            object file = System.IO.Path.GetFullPath(@"Список участников.docx");
            var wordApp = new Application();
            var wordDoc = wordApp.Documents.Open(ref file);

            object oMissing = System.Reflection.Missing.Value;
            wordDoc.Paragraphs[Test_number].Range.Text = Test_number + ". " + Full_name + " " + Age + " " + Sex + " " + DateTime.Now.ToString("dd.MM.yyyy") + " - " + ItogIndex() + "б." + " (" + level() + ")";
            wordDoc.Paragraphs.Add(ref oMissing);

            wordDoc.Save();
            wordDoc.Close();
            wordApp.Quit();
        }

        public void data_word_template()
        {
            string[] data = new string[6];
            string path = System.IO.Path.GetFullPath(@"Сертификат.docx");
            Application word = new Application();
            Document docW = word.Documents.Open(path);
            Bookmarks wBookmarks = docW.Bookmarks;
            Range wRange;

            data[1] = DateTime.Now.ToString("dd.MM.yyyy");
            data[0] = resultSecurity().ToString() + "%"; ;
            data[3] = resultConsume().ToString() + "%"; ;
            data[2] = resultCompetence().ToString() + "%"; ;
            data[4] = level() + ", Ваш ИЦГН - " + Math.Round(ItogIndex(), 2).ToString() + " балла(ов)";
            data[5] = Full_name;

            int i = 0;
            foreach (Bookmark mark in wBookmarks)
            {
                wRange = mark.Range;
                wRange.Text = data[i];
                i++;
            }

            var pathPDF = System.IO.Path.GetFullPath(@"Сертификаты\Сертификат - " + Test_number + ".pdf");
            docW.ExportAsFixedFormat(pathPDF, WdExportFormat.wdExportFormatPDF);
            MessageBox.Show("Необходимо включить принтер", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            word.Quit(false);

            try
            {
                Process p = new Process();
                p.StartInfo = new ProcessStartInfo()
                {
                    CreateNoWindow = true,
                    Verb = "print",
                    FileName = System.IO.Path.GetFullPath(@"Сертификаты\Сертификат - " + Test_number + ".pdf"),
                };
                p.Start();
            }
            catch
            {

            }
        }

        public void print_the_data()
        {
            string[] data = new string[19];
            string path = System.IO.Path.GetFullPath(@"Результаты.docx");
            Application word = new Application();
            Document docW = word.Documents.Open(path);
            docW.Activate();
            Word.Bookmarks wBookmarks = docW.Bookmarks;
            Word.Range wRange;

            data[0] = Full_name + " " + Age + " " + Sex + " " + DateTime.Now.ToString("dd.MM.yyyy") + " г.";
            data[18] = "Итого: безопасность - " + resultSecurity() + "%, потребление - " + resultConsume() + "%, компетенции - " + resultCompetence() + "%" + " (" + level() + "). " + "Время прохождения: 00:" + Transit_Time;

            char[] separators = new char[] { ';' };
            string[] answers_u = Answers_u.Split(separators, StringSplitOptions.RemoveEmptyEntries);

            //Question 1
            if (answers_u[0] == "1")
            {
                data[1] = "Вы ответили: Фамилия, имя, отчество " + AnswersSecurity[0] + "б.";
            }
            else if (answers_u[0] == "2")
            {
                data[1] = "Вы ответили: ДАта и место рождения " + AnswersSecurity[0] + "б.";
            }
            else if (answers_u[0] == "3")
            {
                data[1] = "Вы ответили: Место учебы " + AnswersSecurity[0] + "б.";
            }
            else
            {
                data[1] = "Вы ответили: Все предложенные варианты " + AnswersSecurity[0] + "б.";
            }

            //Question 2 
            if (answers_u[1] == "1")
            {
                data[2] = "Вы ответили: Коля поручился за своего друга, поэтому можно спокойно передать ему пароль " + AnswersSecurity[1] + "б.";
            }
            else if (answers_u[1] == "2")
            {
                data[2] = "Вы ответили: Нет ничего страшного в том, чтобы сообщить пароль другому игроку — это всего лишь игра " + AnswersSecurity[1] + "б.";
            }
            else if (answers_u[1] == "3")
            {
                data[2] = "Вы ответили: Следует отказаться от Колиного предложения, поскольку пользовательское соглашение запрещает игрокам передавать пароль третьим лицам " + AnswersSecurity[1] + "б.";
            }
            else
            {
                data[2] = "Вы ответили: Тане нужно собрать максимум информации о Колином друге, а потом принять окончательное решение " + AnswersSecurity[1] + "б.";
            }

            //Question 3
            if (answers_u[2] == "1")
            {
                data[3] = "Вы ответили: SupermanVasya2005 " + AnswersSecurity[2] + "б.";
            }
            else if (answers_u[2] == "2")
            {
                data[3] = "Вы ответили: QwErTy123456 " + AnswersSecurity[2] + "б.";
            }
            else if (answers_u[2] == "3")
            {
                data[3] = "Вы ответили: Q1jk45)@daJ! " + AnswersSecurity[2] + "б.";
            }
            else
            {
                data[3] = "Вы ответили: M@$h@2oo! " + AnswersSecurity[2] + "б.";
            }

            //Question 4
            if (answers_u[3] == "1")
            {
                data[4] = "Вы ответили: Пройти по ссылке, указанной в письме, и сменить пароль " + AnswersSecurity[2] + "б.";
            }
            else if (answers_u[3] == "2")
            {
                data[4] = "Вы ответили: Проигнорировать письмо и добавить его в спам " + AnswersSecurity[2] + "б.";
            }
            else if (answers_u[3] == "3")
            {
                data[4] = "Вы ответили: Написать в ответ гневное письмо с критикой работы социальной сети " + AnswersSecurity[2] + "б.";
            }
            else if (answers_u[3] == "4")
            {
                data[4] = "Вы ответили: Самостоятельно зайти в свой аккаунт социальной сети и сменить пароль " + AnswersSecurity[2] + "б.";
            }
            else
            {
                data[4] = "Вы ответили: Ответить на это письмо и уточнить информацию " + AnswersSecurity[2] + "б.";
            }

            //Question 5
            if (answers_u[4] == "1")
            {
                data[5] = "Вы ответили: Любое групповое фото, на котором есть изображение данного пользователя " + AnswersSecurity[3] + "б.";
            }
            else if (answers_u[4] == "2")
            {
                data[5] = "Вы ответили: Номер паспорта или любого другого официального документа пользователя " + AnswersSecurity[3] + "б.";
            }
            else
            {
                data[5] = "Вы ответили: Любая персональная информация должна быть удалена из интернета по запросу пользователя " + AnswersSecurity[3] + "б.";
            }

            //Question 6
            if (answers_u[5] == "1")
            {
                data[6] = "Вы ответили: Да " + AnswersConsume[0] + "б.";
            }
            else
            {
                data[6] = "Вы ответили: Нет " + AnswersConsume[0] + "б.";
            }

            //Question 7
            if (answers_u[6] == "1")
            {
                data[7] = "Вы ответили: Регистрация прав сайта " + AnswersConsume[1] + "б.";
            }
            else if (answers_u[6] == "2")
            {
                data[7] = "Вы ответили: Все права защищены " + AnswersConsume[1] + "б.";
            }
            else
            {
                data[7] = "Вы ответили: Копировать разрешено " + AnswersConsume[1] + "б.";
            }

            //Question 8
            data[8] = "Вы ответили: " + answers_u[7] + " " + AnswersSecurity[2] + "б.";

            //Question 9
            if (answers_u[8] == "123")
            {
                data[9] = "Вы ответили: Социальные сети: ВКонтакте, Одноклассники, facebook";
                data[10] = "Мессенджеры: Telegram, WhatsApp, Viber";
                data[11] = "Не относится к вышеперечисленным: GoogleDrive, Skype, Pinterest " + AnswersConsume[3] + "б.";
            }
            else if (answers_u[8] == "132")
            {
                data[9] = "Вы ответили: Социальные сети: ВКонтакте, Одноклассники, facebook";
                data[10] = "Не относится к вышеперечисленным: GoogleDrive, Skype, Pinterest";
                data[11] = "Мессенджеры: Telegram, WhatsApp, Viber " + AnswersConsume[3] + "б.";
            }
            else if (answers_u[8] == "213")
            {
                data[9] = "Вы ответили: Мессенджеры: Telegram, WhatsApp, Viber";
                data[10] = "Социальные сети: ВКонтакте, Одноклассники, facebook";
                data[11] = "Не относится к вышеперечисленным: GoogleDrive, Skype, Pinterest " + AnswersConsume[3] + "б.";
            }
            else if (answers_u[8] == "231")
            {
                data[9] = "Вы ответили: Мессенджеры: Telegram, WhatsApp, Viber";
                data[10] = "Не относится к вышеперечисленным: GoogleDrive, Skype, Pinterest";
                data[11] = "Социальные сети: ВКонтакте, Одноклассники, facebook " + AnswersConsume[3] + "б.";
            }
            else if (answers_u[8] == "312")
            {
                data[9] = "Вы ответили: Не относится к вышеперечисленным: GoogleDrive, Skype, Pinterest";
                data[10] = "Социальные сети: ВКонтакте, Одноклассники, facebook";
                data[11] = "Мессенджеры: Telegram, WhatsApp, Viber " + AnswersConsume[3] + "б.";
            }
            else //321
            {
                data[9] = "Вы ответили: Не относится к вышеперечисленным: GoogleDrive, Skype, Pinterest";
                data[10] = "Мессенджеры: Telegram, WhatsApp, Viber";
                data[11] = "Социальные сети: ВКонтакте, Одноклассники, facebook " + AnswersConsume[3] + "б.";
            }

            //Question 10
            if (answers_u[9] == "1")
            {
                data[12] = "Вы ответили: Счет для взаимодействия с которым нужно более одного ключа " + AnswersCompetence[0] + "б.";
            }
            else if (answers_u[9] == "2")
            {
                data[12] = "Вы ответили: Интегральная схема, разработанная для выполнения определенной задачи " + AnswersCompetence[0] + "б.";
            }
            else
            {
                data[12] = "Вы ответили: Выстроенная по определенным правилам непрерывная последовательная цепочка блоков (связный список), содержащих информацию " + AnswersCompetence[0] + "б.";
            }

            //Question 11
            if (answers_u[10] == "1")
            {
                data[13] = "Вы ответили: Да " + AnswersCompetence[1] + "б.";
            }
            else
            {
                data[13] = "Вы ответили: Нет " + AnswersCompetence[1] + "б.";
            }

            //Question 12
            if (answers_u[11] == "1")
            {
                data[14] = "Вы ответили: Для того, чтобы знать, как зовут пользователя " + AnswersCompetence[2] + "б.";
            }
            else if (answers_u[11] == "2")
            {
                data[14] = "Вы ответили: Для эстетического вида " + AnswersCompetence[2] + "б.";
            }
            else
            {
                data[14] = "Вы ответили: Для поиска компьютера в сети " + AnswersCompetence[2] + "б.";
            }

            //Question 13
            if (answers_u[12] == "1")
            {
                data[15] = "Вы ответили: «Все пошло не так» " + AnswersCompetence[3] + "б.";
            }
            else if (answers_u[12] == "2")
            {
                data[15] = "Вы ответили: «Откуда опять взялась эта ошибка в программе» " + AnswersCompetence[3] + "б.";
            }
            else if (answers_u[12] == "3")
            {
                data[15] = "Вы ответили: «Программа не работает, потому что требуется предоплата» " + AnswersCompetence[3] + "б.";
            }
            else
            {
                data[15] = "Вы ответили: «Это не ошибка. Так и было задумано» " + AnswersCompetence[3] + "б.";
            }

            //Question 14
            if (answers_u[13] == "1")
            {
                data[16] = "Вы ответили: Валюта, у которой засекречен источник ее выпуска " + AnswersCompetence[4] + "б.";
            }
            else if (answers_u[13] == "2")
            {
                data[16] = "Вы ответили: Электронная валюта, у которой нет администратора – ее стоимость не устанавливается и не гарантируется ни одним государством " + AnswersCompetence[4] + "б.";
            }
            else if (answers_u[13] == "3")
            {
                data[16] = "Вы ответили: Валюта, которую выпускает банк только в электронном виде " + AnswersCompetence[4] + "б.";
            }
            else
            {
                data[16] = "Вы ответили: Электронная валюта, все сделки с которой проводятся скрытно " + AnswersCompetence[4] + "б.";
            }

            //Question 15
            if (answers_u[14] == "1")
            {
                data[17] = "Вы ответили: Оплачивать услуги и переводить на банковские счета, но только частным лицам " + AnswersConsume[4] + "б.";
            }
            else if (answers_u[14] == "2")
            {
                data[17] = "Вы ответили: Отправлять, получать и хранить " + AnswersConsume[4] + "б.";
            }
            else if (answers_u[14] == "2")
            {
                data[17] = "Вы ответили: Продавать и переводить в другие валюты, но только не в рубли " + AnswersConsume[4] + "б.";
            }
            else
            {
                data[17] = "Вы ответили: Законом не запрещено только говорить о них " + AnswersConsume[4] + "б.";
            }


            int j = 0;
            foreach (Word.Bookmark mark in wBookmarks)
            {
                wRange = mark.Range;
                wRange.Text = data[j];
                j++;
            }

            var pathPDF = System.IO.Path.GetFullPath(@"Результаты\Результаты - " + Test_number + ".pdf");
            docW.ExportAsFixedFormat(pathPDF, WdExportFormat.wdExportFormatPDF);
            MessageBox.Show("Необходимо включить принтер", "Внимание");
            word.Quit(false);

            try
            {
                Process p = new Process();
                p.StartInfo = new ProcessStartInfo()
                {
                    CreateNoWindow = true,
                    Verb = "print",
                    FileName = System.IO.Path.GetFullPath(@"Результаты\Результаты - " + Test_number + ".pdf"),
                };
                p.Start();
            }
            catch
            {

            }
        }
    }
}
