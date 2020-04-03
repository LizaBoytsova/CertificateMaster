using System;
using System.IO;
using System.Drawing.Printing;
using System.Windows.Forms;
using System.Xml.Serialization;

namespace CertificateMaster
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            try
            {
                // Показать последние введённые данные
                ShowLastData(DeSerializeNote(Way.LastData));
            }
            catch
            {
                MessageBox.Show("Не удалось загрузить последние данные");
                AllData data = new AllData();
                data.CountCert = 1;
                rBtCountCert1.Checked = true;
                PrinterSettings.StringCollection sc = PrinterSettings.InstalledPrinters;
                data.PrintName = sc[1];
            }
            this.Text = "Мастер удостоверений вер. " + System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString();
            CheckDataFolder();
        }

        /// <summary>
        /// Проверка на наличие папки с предыдущими данными
        /// </summary>
        public void CheckDataFolder()
        {
            string MyDir = Way.MainWay + "Данные";

            if (!Directory.Exists(MyDir))
                Directory.CreateDirectory(MyDir);
        }

        /// <summary>
        /// Вызов запроса на очистку всех полей
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnClean_Click(object sender, EventArgs e)
        {
            var result = MessageBox.Show("Вы уверены, что хотите очистить все поля?", "Очистить поля", MessageBoxButtons.YesNo);
            if (result == DialogResult.Yes)
                CleanAll();
        }

        /// <summary>
        /// Очистка всех полей
        /// </summary>
        private void CleanAll()
        {
            CleanTab1();
            CleanTab2();
        }

        /// <summary>
        /// Очистка вкладки 1
        /// </summary>
        private void CleanTab1()
        {
            txtCertNumTab1.Value = 1;
            txtFIOTab1.Text = "";
            txtPost1Tab1.Text = "";
            txtPost2Tab1.Text = "";
            txtWorkplace1Tab1.Text = "";
            txtWorkplace2Tab1.Text = "";
            dateStartLearningTab1.Value = DateTime.Today;
            dateEndLearningTab1.Value = DateTime.Today;
            txtLearningForWhoTab1.Text = "";
            txtProtocolNumberTab1.Value = 1;
            dateProtocolTab1.Value = DateTime.Today;
            dateEndCertificateTab1.Value = DateTime.Today;
        }

        /// <summary>
        /// Очистка вкладки 2
        /// </summary>
        private void CleanTab2()
        {
            txtCertNumTab2.Value = 1;
            txtFIOTab2.Text = "";
            txtPost1Tab2.Text = "";
            txtPost2Tab2.Text = "";
            txtWorkplace1Tab2.Text = "";
            txtWorkplace2Tab2.Text = "";
            dateStartLearningTab2.Value = DateTime.Today;
            dateEndLearningTab2.Value = DateTime.Today;
            txtLearningForWhoTab2.Text = "";
            txtProtocolNumberTab2.Value = 1;
            dateProtocolTab2.Value = DateTime.Today;
            dateEndCertificateTab2.Value = DateTime.Today;
        }

        /// <summary>
        /// Получение информации из полей формы и сохранение в класс
        /// </summary>
        private AllData GetDataFromFields()
        {
            AllData Data = new AllData();

            if (rBtCountCert1.Checked)
                Data.CountCert = 1;
            if (rBtCountCert2.Checked)
                Data.CountCert = 2;

            Data.justPrint = chBxJustPrint.Checked;

            Data.CertNumTab1 = (int)txtCertNumTab1.Value;
            Data.FIOTab1 = txtFIOTab1.Text;
            Data.Post1Tab1 = txtPost1Tab1.Text;
            Data.Post2Tab1 = txtPost2Tab1.Text;
            Data.Workplace1Tab1 = txtWorkplace1Tab1.Text;
            Data.Workplace2Tab1 = txtWorkplace2Tab1.Text;
            Data.dateStartLearningTab1 = dateStartLearningTab1.Value;
            Data.dateEndLearningTab1 = dateEndLearningTab1.Value;
            Data.LearningForWhoTab1 = txtLearningForWhoTab1.Text;
            Data.ProtocolNumberTab1 = (int)txtProtocolNumberTab1.Value;
            Data.dateProtocolTab1 = dateProtocolTab1.Value;
            Data.dateEndCertificateTab1 = dateEndCertificateTab1.Value;

            Data.CertNumTab2 = (int)txtCertNumTab2.Value;
            Data.FIOTab2 = txtFIOTab2.Text;
            Data.Post1Tab2 = txtPost1Tab2.Text;
            Data.Post2Tab2 = txtPost2Tab2.Text;
            Data.Workplace1Tab2 = txtWorkplace1Tab2.Text;
            Data.Workplace2Tab2 = txtWorkplace2Tab2.Text;
            Data.dateStartLearningTab2 = dateStartLearningTab2.Value;
            Data.dateEndLearningTab2 = dateEndLearningTab2.Value;
            Data.LearningForWhoTab2 = txtLearningForWhoTab2.Text;
            Data.ProtocolNumberTab2 = (int)txtProtocolNumberTab2.Value;
            Data.dateProtocolTab2 = dateProtocolTab2.Value;
            Data.dateEndCertificateTab2 = dateEndCertificateTab2.Value;


            return Data;
        }

        /// <summary>
        /// Cохранение информации из полей формы в файл
        /// </summary>
        private void InsertDataForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            try
            {
                SerializeNote(GetDataFromFields(), Way.LastData);
            }
            catch
            {
                MessageBox.Show("Не удалось сохранить введеные данные");
            }
        }

        /// <summary>
        /// Сохранение последних введённых данных в файл
        /// </summary>
        /// <param name="data"></param>
        /// <param name="path"></param>
        private void SerializeNote(AllData data, string path)
        {
            XmlSerializer ser = new XmlSerializer(typeof(AllData));
            StreamWriter sw = new StreamWriter(path);
            ser.Serialize(sw, data);
            sw.Close();
        }

        /// <summary>
        /// Загрузка предыдущих введенных данных из файла
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        private AllData DeSerializeNote(string path)
        {
            XmlSerializer ser = new XmlSerializer(typeof(AllData));
            StreamReader streamR = new StreamReader(path);
            AllData data = (AllData)ser.Deserialize(streamR);
            streamR.Close();

            return data;
        }

        /// <summary>
        /// Отображение последней введенной информации в окне ввода данных
        /// </summary>
        /// <param name="Data"></param>
        private void ShowLastData(AllData Data)
        {
            if (Data.CountCert == 2)
                rBtCountCert2.Checked = true;
            else
                rBtCountCert1.Checked = true;
            
            chBxJustPrint.Checked = Data.justPrint;

            txtCertNumTab1.Value = Data.CertNumTab1;
            txtFIOTab1.Text = Data.FIOTab1;
            txtPost1Tab1.Text = Data.Post1Tab1;
            txtPost2Tab1.Text = Data.Post2Tab1;
            txtWorkplace1Tab1.Text = Data.Workplace1Tab1;
            txtWorkplace2Tab1.Text = Data.Workplace2Tab1;
            dateStartLearningTab1.Value = Data.dateStartLearningTab1;
            dateEndLearningTab1.Value = Data.dateEndLearningTab1;
            txtLearningForWhoTab1.Text = Data.LearningForWhoTab1;
            txtProtocolNumberTab1.Value = Data.ProtocolNumberTab1;
            dateProtocolTab1.Value = Data.dateProtocolTab1;
            dateEndCertificateTab1.Value = Data.dateEndCertificateTab1;


            txtCertNumTab2.Value = Data.CertNumTab2;
            txtFIOTab2.Text = Data.FIOTab2;
            txtPost1Tab2.Text = Data.Post1Tab2;
            txtPost2Tab2.Text = Data.Post2Tab2;
            txtWorkplace1Tab2.Text = Data.Workplace1Tab2;
            txtWorkplace2Tab2.Text = Data.Workplace2Tab2;
            dateStartLearningTab2.Value = Data.dateStartLearningTab2;
            dateEndLearningTab2.Value = Data.dateEndLearningTab2;
            txtLearningForWhoTab2.Text = Data.LearningForWhoTab2;
            txtProtocolNumberTab2.Value = Data.ProtocolNumberTab2;
            dateProtocolTab2.Value = Data.dateProtocolTab2;
            dateEndCertificateTab2.Value = Data.dateEndCertificateTab2;

        }

        private void btnClean_Click(object sender, EventArgs e)
        {
            var result = MessageBox.Show("Вы уверены, что хотите очистить все поля?\nБудут удалены данные с обоих вкладок!", "Очистить поля", MessageBoxButtons.YesNo);
            if (result == DialogResult.Yes)
                CleanAll();
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            try
            {
                SerializeNote(GetDataFromFields(), Way.LastData);
            }
            catch
            {
                MessageBox.Show("Не удалось сохранить введеные данные");
            }
        }

        private void btGO_Click(object sender, EventArgs e)
        {
            if (AllDataInsert())
            {
                AllData data = GetDataFromFields();
                SerializeNote(data, Way.LastData);

                CreateCert(data);
                
            }
        }

        /// <summary>
        /// Заполнение одного документа СЗ
        /// </summary>
        /// <param name="Now"></param>
        /// <returns></returns>
        private bool CreateCert(AllData data)
        {
            WordDocument wordDoc = new WordDocument(Way.Certificate1);
            if (rBtCountCert2.Checked)
                wordDoc = new WordDocument(Way.Certificate2);

            try
            {
                // замена первого сертификата
                wordDoc.ReplaceAllStrings("номер", data.CertNumTab1.ToString());
                wordDoc.ReplaceAllStrings("Фио", data.FIOTab1);
                wordDoc.ReplaceAllStrings("11", data.dateStartLearningTab1.Day.ToString());
                wordDoc.ReplaceAllStrings("месяц1", MonthToString(data.dateStartLearningTab1));
                wordDoc.ReplaceAllStrings("22", data.dateEndLearningTab1.Day.ToString());
                wordDoc.ReplaceAllStrings("месяц2", MonthToString(data.dateEndLearningTab1));
                wordDoc.ReplaceAllStrings("кого", data.LearningForWhoTab1);
                wordDoc.ReplaceAllStrings("нпр", data.ProtocolNumberTab1.ToString());
                wordDoc.ReplaceAllStrings("33", data.dateProtocolTab1.Day.ToString());
                wordDoc.ReplaceAllStrings("месяц3", MonthToString(data.dateProtocolTab1));
                wordDoc.ReplaceAllStrings("44", data.dateEndCertificateTab1.Day.ToString());
                wordDoc.ReplaceAllStrings("месяц4", MonthToString(data.dateEndCertificateTab1));
                wordDoc.ReplaceAllStrings("гг", data.dateEndCertificateTab1.Year.ToString().Substring(2));

                string max = data.Post1Tab1;
                if (data.Post2Tab1.Length > data.Post1Tab1.Length) max = data.Post2Tab1;

                ChangeFontSize(wordDoc, "Должность1", max);
                wordDoc.ReplaceAllStrings("Должность1", data.Post1Tab1);

                ChangeFontSize(wordDoc, "Должность2", max);
                wordDoc.ReplaceAllStrings("Должность2", data.Post2Tab1);

                max = data.Workplace1Tab1;
                if (data.Workplace2Tab1.Length > data.Workplace1Tab1.Length) max = data.Workplace2Tab1;

                ChangeFontSize(wordDoc, "Место работы1", max);
                wordDoc.ReplaceAllStrings("Место работы1", data.Workplace1Tab1);

                ChangeFontSize(wordDoc, "Место работы2", max);
                wordDoc.ReplaceAllStrings("Место работы2", data.Workplace2Tab1);

                // замена второго сертификата
                wordDoc.ReplaceAllStrings("номер", data.CertNumTab1.ToString());
                wordDoc.ReplaceAllStrings("Фио", data.FIOTab1);
                wordDoc.ReplaceAllStrings("11", data.dateStartLearningTab1.Day.ToString());
                wordDoc.ReplaceAllStrings("месяц1", MonthToString(data.dateStartLearningTab1));
                wordDoc.ReplaceAllStrings("22", data.dateEndLearningTab1.Day.ToString());
                wordDoc.ReplaceAllStrings("месяц2", MonthToString(data.dateEndLearningTab1));
                wordDoc.ReplaceAllStrings("кого", data.LearningForWhoTab1);
                wordDoc.ReplaceAllStrings("нпр", data.ProtocolNumberTab1.ToString());
                wordDoc.ReplaceAllStrings("33", data.dateProtocolTab1.Day.ToString());
                wordDoc.ReplaceAllStrings("месяц3", MonthToString(data.dateProtocolTab1));
                wordDoc.ReplaceAllStrings("44", data.dateEndCertificateTab1.Day.ToString());
                wordDoc.ReplaceAllStrings("месяц4", MonthToString(data.dateEndCertificateTab1));
                wordDoc.ReplaceAllStrings("гг", data.dateEndCertificateTab1.Year.ToString().Substring(2));


                string max2 = data.Post1Tab1;
                if (data.Post2Tab1.Length > data.Post1Tab1.Length) max2 = data.Post2Tab1;

                ChangeFontSize(wordDoc, "Должность1", max2);
                wordDoc.ReplaceAllStrings("Должность1", data.Post1Tab1);

                ChangeFontSize(wordDoc, "Должность2", max2);
                wordDoc.ReplaceAllStrings("Должность2", data.Post2Tab1);

                max2 = data.Workplace1Tab1;
                if (data.Workplace2Tab1.Length > data.Workplace1Tab1.Length) max2 = data.Workplace2Tab1;

                ChangeFontSize(wordDoc, "Место работы1", max2);
                wordDoc.ReplaceAllStrings("Место работы1", data.Workplace1Tab1);

                ChangeFontSize(wordDoc, "Место работы2", max2);
                wordDoc.ReplaceAllStrings("Место работы2", data.Workplace2Tab1);


                //печать или отображение документа
                if (chBxJustPrint.Checked)
                {
                    wordDoc.DocPrint();
                    wordDoc.Close();
                }
                else
                    wordDoc.Visible = true;

            }
            catch (Exception error)
            {
                if (wordDoc != null) { wordDoc.Close(); }
                MessageBox.Show("Ошибка при вставке текста в документ  Word. Подробности " + error.Message);
                return false;
            }

            return true;
        }


        private void ChangeFontSize(WordDocument wordDoc, string change, string str)
        {
            wordDoc.SetSelectionToText(change);
            wordDoc.Selection.Italic = true;
            wordDoc.Selection.Aligment = TextAligment.Center;
            if (str.Length >= 42 && str.Length <= 45)
                wordDoc.Selection.FontSize = 10;
            else if (str.Length >= 46 && str.Length <= 51)
                wordDoc.Selection.FontSize = 9;
            else if (str.Length >= 52 && str.Length <= 58)
                wordDoc.Selection.FontSize = 8;
            else if (str.Length >= 59 && str.Length <= 66)
                wordDoc.Selection.FontSize = 7;
            else if (str.Length >= 67 && str.Length <= 78)
                wordDoc.Selection.FontSize = 6;
            else if (str.Length >= 79 && str.Length <= 94)
                wordDoc.Selection.FontSize = 5;
            else if (str.Length >= 95 && str.Length <= 118)
                wordDoc.Selection.FontSize = 4;
            else if (str.Length >= 119)
                wordDoc.Selection.FontSize = 3;
        }

        /// <summary>
        /// Формирования строки даты на русском языке
        /// </summary>
        /// <param name="Date"></param>
        /// <returns></returns>
        private string MonthToString(DateTime Date)
        {
            string s = "";
            switch (Date.Month)
            {
                case 1:
                    s = "января ";
                    break;
                case 2:
                    s = "февраля ";
                    break;
                case 3:
                    s = "марта ";
                    break;
                case 4:
                    s = "апреля ";
                    break;
                case 5:
                    s = "мая ";
                    break;
                case 6:
                    s = "июня ";
                    break;
                case 7:
                    s = "июля ";
                    break;
                case 8:
                    s = "августа ";
                    break;
                case 9:
                    s = "сентября ";
                    break;
                case 10:
                    s = "октября ";
                    break;
                case 11:
                    s = "ноября ";
                    break;
                case 12:
                    s = "декабря ";
                    break;
                default:
                    break;

            }
            s += Date.Year.ToString();
            return s;
        }



        /// <summary>
        /// Проверка ввода корректных данных
        /// </summary>
        private bool AllDataInsert()
        {

            int i = 0; // индекс количества ошибок
            string str = "Не указаны следующие параметры: \n"; // строка, содержащая описание ошибок

            //
            if (txtFIOTab1.Text == "")
            {
                str += "ФИО кому выдано (Сертификат №1)\n"; i++;
            }

            //
            if (txtPost1Tab1.Text == "")
            {
                str += "Должность (Сертификат №1)\n"; i++;
            }

            //
            if (txtWorkplace1Tab1.Text == "")
            {
                str += "Место работы (Сертификат №1)\n"; i++;
            }

            //
            if (txtLearningForWhoTab1.Text == "")
            {
                str += "Для каких лиц обучение (Сертификат №1)\n"; i++;
            }

            if (txtFIOTab2.Text == "")
            {
                str += "ФИО кому выдано (Сертификат №2)\n"; i++;
            }

            //
            if (txtPost1Tab2.Text == "")
            {
                str += "Должность (Сертификат №2)\n"; i++;
            }

            //
            if (txtWorkplace1Tab2.Text == "")
            {
                str += "Место работы (Сертификат №2)\n"; i++;
            }

            //
            if (txtLearningForWhoTab2.Text == "")
            {
                str += "Для каких лиц обучение (Сертификат №2)\n"; i++;
            }

            //
            if (i != 0)
            {
                var result = MessageBox.Show(str + "\nВсё равно сформировать?", "Ошибка ввода", MessageBoxButtons.YesNo);
                if (result == DialogResult.Yes)
                    return true;
                else
                    return false;
            }
            else return true;
        }

        private void справкаToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Программа Мастер удостоверений \nВерсия " + System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString() + "\n2018 год\n\nДанная программа предназначена для формирования удостоверения об обучении пожарно-техническому минимуму\n\nРазработчик Бойцова Е.С.", "О программе",
                MessageBoxButtons.OK, MessageBoxIcon.None);
        }

        private void dateProtocolTab1_ValueChanged(object sender, EventArgs e)
        {
            dateEndCertificateTab1.Value = dateProtocolTab1.Value;
            dateEndCertificateTab1.Value = dateProtocolTab1.Value.AddYears(3);
            dateEndCertificateTab1.Value = dateEndCertificateTab1.Value.AddDays(1);
        }

        private void dateProtocolTab2_ValueChanged(object sender, EventArgs e)
        {
            dateEndCertificateTab2.Value = dateProtocolTab2.Value;
            dateEndCertificateTab2.Value = dateProtocolTab2.Value.AddYears(3);
            dateEndCertificateTab2.Value = dateEndCertificateTab2.Value.AddDays(1);
        }

        /// <summary>
        /// Открытие папки с шаблонами
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void открытьПапкуСШаблонамиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var result = MessageBox.Show("Внимание! Редактирование данных файлов может привести к ошибкам выполнения программы!", "Предупреждение", MessageBoxButtons.OKCancel);
            if (result == DialogResult.OK)
                System.Diagnostics.Process.Start("explorer.exe", Way.MainWay + "Шаблоны");
        }

        private void rBtCountCert1_CheckedChanged(object sender, EventArgs e)
        {
            if (rBtCountCert1.Checked)
            {
                foreach (Control c in tabCert2.Controls)
                    c.Enabled = false;
            }
        }

        private void rBtCountCert2_CheckedChanged(object sender, EventArgs e)
        {
            if (rBtCountCert2.Checked)
            {
                foreach (Control c in tabCert2.Controls)
                    c.Enabled = true;
            }
        }

        private void btCleanTab2_Click(object sender, EventArgs e)
        {
            var result = MessageBox.Show("Вы уверены, что хотите очистить все поля данной вкладки?", "Очистить поля", MessageBoxButtons.YesNo);
            if (result == DialogResult.Yes)
            CleanTab2();
        }

        private void btCleanTab1_Click(object sender, EventArgs e)
        {
            var result = MessageBox.Show("Вы уверены, что хотите очистить все поля данной вкладки?", "Очистить поля", MessageBoxButtons.YesNo);
            if (result == DialogResult.Yes)
                CleanTab1();
        }

    }
}
