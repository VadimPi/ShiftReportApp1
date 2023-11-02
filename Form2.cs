using Npgsql;
using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.Common;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static Microsoft.EntityFrameworkCore.DbLoggerCategory.Database;
using DocumentFormat.OpenXml.Spreadsheet;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Button;

namespace ShiftReportApp1
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();

            checkBox17.Checked = true;
            checkBox18.Checked = true;
            checkBox15.Checked = true;
            checkBox14.Checked = true;
            checkBox13.Checked = true;
            checkBox12.Checked = true;
            checkBox11.Checked = true;

            dateTimePicker1.Enabled = false;
            dateTimePicker2.Enabled = false;
            numericUpDown1.Enabled = false;
            numericUpDown2.Enabled = false;
            numericUpDown3.Enabled = false;
            groupBox4.Enabled = false;
            groupBox3.Enabled = false;

        }

        private void ExecuteQuery(string query)
        {
            try
            {
                ProjectLogger.LogDebug("Начало ExecuteQuery");
                checkedListBox1.Items.Clear();

                DataBaseConnection dbConnection = new DataBaseConnection();
                NpgsqlConnection connection = dbConnection.GetConnection();

                connection.Open();
                using (NpgsqlCommand cmd = new NpgsqlCommand(query, connection))
                {
                    (DateTime varDate1, DateTime varDate2) = GetDateRange();

                    cmd.Parameters.AddWithValue("@varDate1", varDate1);
                    cmd.Parameters.AddWithValue("@varDate2", varDate2);

                    using (NpgsqlDataReader reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            string productName = $"{reader["product_name"]}";
                            string depth = $"{reader["prod_depth"]}";
                            string length = $"{reader["length"]}";
                            string width = $"{reader["width"]}";
                            bool unspecified = (bool)reader["unspecified"];

                            string outputString;

                            if (unspecified)
                            {
                                // Последние три символа productName
                                string lastThreeChars = productName.Substring(productName.Length - 3);
                                outputString = $"{productName} {length}x{width}x{depth}({lastThreeChars})";
                            }
                            else
                            {
                                outputString = $"{productName} {length}x{width}x{depth}";
                            }

                            checkedListBox1.Items.Add(outputString);
                        }
                    }
                }
                ProjectLogger.LogDebug("Конец ExecuteQuery");
            }
            catch (Exception ex)
            {
                // Обработка ошибки
                MessageBox.Show($"An error occurred: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                ProjectLogger.LogException("Ошибка в ExecuteQuery", ex);
            }
        }
        private (DateTime, DateTime) GetDateRange()
        {
            DateTime now = DateTime.Now;
            int currentYear = now.Year;
            DateTime varDate1 = DateTime.MinValue;
            DateTime varDate2 = DateTime.MinValue;

            if (dateTimePicker1.Enabled)
            {
                varDate1 = dateTimePicker1.Value;
                varDate2 = dateTimePicker2.Value;
            }
            else
            {
                if (numericUpDown1.Enabled)
                {
                    int selectedWeek = (int)numericUpDown1.Value;
                    // Получаем первый день года
                    DateTime firstDayOfYear = new DateTime(currentYear, 1, 1);
                    // Получаем первый день текущей недели
                    DateTime firstDayOfSelectedWeek = firstDayOfYear.AddDays((selectedWeek - 1) * 7);
                    // Понедельник
                    varDate1 = firstDayOfSelectedWeek.AddDays(DayOfWeek.Monday - firstDayOfSelectedWeek.DayOfWeek);
                    // Воскресенье
                    varDate2 = varDate1.AddDays(6);
                }
                else if (numericUpDown2.Enabled)
                {
                    int selectedMonth = (int)numericUpDown2.Value;
                    varDate1 = new DateTime(currentYear, selectedMonth, 1);
                    varDate2 = varDate1.AddMonths(1).AddDays(-1);
                }
                else if (numericUpDown3.Enabled)
                {
                    int selectedYear = (int)numericUpDown3.Value;
                    varDate1 = new DateTime(selectedYear, 1, 1);
                    varDate2 = new DateTime(selectedYear, 12, 31);
                }
                else if (comboBox1.SelectedIndex == 0)
                {
                    if (9 <= now.Hour && now.Hour < 21)
                    {
                        varDate1 = now.Date.AddDays(-1).AddHours(20);
                        varDate2 = varDate1.AddHours(12);
                    }
                    if (now.Hour >= 21 && now.Hour <= 23)
                    {
                        varDate1 = now.Date.AddHours(8);
                        varDate2 = varDate1.AddHours(12);
                    }
                    if (0 <= now.Hour && now.Hour < 9)
                    {
                        varDate1 = now.Date.AddDays(-1).AddHours(8);
                        varDate2 = varDate1.AddHours(12);
                    }
                }
                else if (comboBox1.SelectedIndex == 1)
                {
                    varDate1 = now.Date.AddDays(-1);
                    varDate2 = varDate1.AddHours(23).AddMinutes(59).AddSeconds(59);
                }
            }
            if (varDate1 > varDate2)
            {
                (varDate1, varDate2) = (varDate2, varDate1);
            }

            return (varDate1, varDate2);
        }


        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void checkBox13_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            groupBox1.Enabled = true;
            groupBox2.Enabled = false;
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            groupBox1.Enabled = false;
            groupBox2.Enabled = true;
        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            // Делаем все элементы неактивными
            dateTimePicker1.Enabled = false;
            dateTimePicker2.Enabled = false;
            numericUpDown1.Enabled = false;
            numericUpDown2.Enabled = false;
            numericUpDown3.Enabled = false;
            groupBox4.Enabled = false;
            groupBox3.Enabled = false;

            // Проверяем выбранный элемент в ComboBox
            switch (comboBox1.SelectedIndex)
            {
                case 2: // С даты по дату
                    dateTimePicker1.Enabled = true;
                    dateTimePicker2.Enabled = true;
                    groupBox4.Enabled = true;
                    groupBox3.Enabled = true;
                    break;
                case 3: // За неделю
                    numericUpDown1.Enabled = true;
                    groupBox4.Enabled = true;
                    groupBox3.Enabled = true;
                    break;
                case 5: // За месяц
                    numericUpDown3.Enabled = true;
                    groupBox4.Enabled = true;
                    groupBox3.Enabled = true;
                    break;
                case 4: // За год
                    numericUpDown2.Enabled = true;
                    groupBox4.Enabled = true;
                    groupBox3.Enabled = true;
                    break;
                default: // За предыдущую смену и За предыдущие сутки
                    break;
            }

            DateTime now = DateTime.Now;

            if (radioButton1.Checked && now.Hour >= 9 && now.Hour < 21 && comboBox1.SelectedIndex == 0)
            {
                ExecuteQuery(Names.Request(1)); // за предыдущую смену после 9-00
            }
            else if (radioButton1.Checked && now.Hour <= 9 && comboBox1.SelectedIndex == 0)
            {
                ExecuteQuery(Names.Request(2)); // за предыдущую смену после 00-00 до 9-00
            }
            else if (radioButton1.Checked && now.Hour >= 21 && comboBox1.SelectedIndex == 0)
            {
                ExecuteQuery(Names.Request(3)); //за предыдущую смену после 21-00 до 23-59
            }
            else if (radioButton1.Checked && comboBox1.SelectedIndex == 1)
            {
                ExecuteQuery(Names.Request(4)); // за предыдущие сутки после 9-00 до 23-59
            }
            else if (radioButton1.Checked && comboBox1.SelectedIndex == 2)
            {
                ExecuteQuery(Names.Request(5)); // с даты по дату
            }
        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void label5_Click(object sender, EventArgs e)
        {

        }

        private void label6_Click(object sender, EventArgs e)
        {

        }

        private void label7_Click(object sender, EventArgs e)
        {

        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void label8_Click(object sender, EventArgs e)
        {

        }

        private void label9_Click(object sender, EventArgs e)
        {

        }

        private void groupBox2_Enter(object sender, EventArgs e)
        {

        }

        private void groupBox5_Enter(object sender, EventArgs e)
        {

        }

        private void groupBox3_Enter(object sender, EventArgs e)
        {

        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            if (radioButton1.Checked)
            {
                ExecuteQuery(Names.Request(5));
            }
        }

        private void dateTimePicker2_ValueChanged(object sender, EventArgs e)
        {
            if (radioButton1.Checked)
            {
                ExecuteQuery(Names.Request(5));
            }
        }

        private void numericUpDown1_ValueChanged(object sender, EventArgs e)
        {
            if (radioButton1.Checked)
            {
                ExecuteQuery(Names.Request(5));
            }
        }

        private void numericUpDown2_ValueChanged(object sender, EventArgs e)
        {
            if (radioButton1.Checked)
            {
                ExecuteQuery(Names.Request(5));
            }
        }

        private void numericUpDown3_ValueChanged(object sender, EventArgs e)
        {
            if (radioButton1.Checked)
            {
                ExecuteQuery(Names.Request(5));
            }
        }

        private void checkedListBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void checkedListBox2_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Копируем все элементы из checkedListBox1 в checkedListBox2
            foreach (var item in checkedListBox1.Items)
            {
                if (!checkedListBox2.Items.Contains(item))
                {
                    checkedListBox2.Items.Add(item);
                }
            }

            for (int i = 0; i < checkedListBox2.Items.Count; i++)
            {
                checkedListBox2.SetItemChecked(i, true);
            }
            // Очищаем checkedListBox1
            checkedListBox1.Items.Clear();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            List<object> CheckedItems = checkedListBox1.CheckedItems.Cast<object>().ToList();

            foreach (var сheckedItem in CheckedItems)
            {
                checkedListBox2.Items.Add(сheckedItem);
                checkedListBox1.Items.Remove(сheckedItem);
            }
            for (int i = 0; i < checkedListBox2.Items.Count; i++)
            {
                checkedListBox2.SetItemChecked(i, true);
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            List<object> CheckedItems = checkedListBox2.CheckedItems.Cast<object>().ToList();

            foreach (var сheckedItem in CheckedItems)
            {
                checkedListBox1.Items.Add(сheckedItem);
                checkedListBox2.Items.Remove(сheckedItem);
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            // Копируем все элементы из checkedListBox2 в checkedListBox1
            foreach (var item in checkedListBox2.Items)
            {
                if (!checkedListBox1.Items.Contains(item))
                {
                    checkedListBox1.Items.Add(item);
                }
            }

            // Очищаем checkedListBox2
            checkedListBox2.Items.Clear();
        }

        private void checkBox16_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox16.Checked == true)
            {
                checkBox17.Checked = false;
                checkBox18.Checked = false;
                checkBox17.Enabled = false;
                checkBox18.Enabled = false;
            }
            else
            {
                checkBox17.Enabled = true;
                checkBox18.Enabled = true;
                checkBox17.Checked = true;
                checkBox18.Checked = true;
            }
        }

        private void checkBox17_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void checkBox18_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void checkBox10_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void checkBox9_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void checkBox8_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void checkBox7_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void checkBox6_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void checkBox15_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void checkBox14_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void checkBox12_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void checkBox11_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void checkBox5_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void checkBox4_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void checkBox19_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void checkBox20_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void button6_Click(object sender, EventArgs e)
        {
            //инициализируем состояние кнопок смен
            int numShift1 = checkBox9.Checked ? 1 : 0;
            int numShift2 = checkBox8.Checked ? 2 : 0;
            int numShift3 = checkBox7.Checked ? 3 : 0;
            int numShift4 = checkBox6.Checked ? 4 : 0;

            string typeStops1 = checkBox14.Checked ? "Технологические" : "_";
            string typeStops2 = checkBox13.Checked ? "Настройки" : "_";
            string typeStops3 = checkBox12.Checked ? "Поломки" : "_";
            string typeStops4 = checkBox11.Checked ? "ППР" : "_";

            // инициализируем даты
            (DateTime varDate1, DateTime varDate2) = GetDateRange();
            DateTime now = DateTime.Now;

            // запишем в список все отмеченные элементы в checkedListBox2
            List<string> prodList = checkedListBox2.CheckedItems.Cast<string>().ToList();


            int queryNumb = GetQueryNumber(now);

            SaAsDi saver = new SaAsDi();  // Создаем экземпляр класса SaAsDi
            saver.SaveExcelFile(queryNumb, prodList, varDate1, varDate2, numShift1, numShift2, numShift3, numShift4,
                typeStops1, typeStops2, typeStops3, typeStops4);       // Вызываем метод сохранения
        }

        private void button5_Click(object sender, EventArgs e)
        {

        }

        private void button7_Click(object sender, EventArgs e)
        {
            Form2 form2 = Application.OpenForms.OfType<Form2>().FirstOrDefault();
            form2.Close();
        }

        private int GetQueryNumber(DateTime now)
        {
            if (radioButton1.Checked)
            {
                if (comboBox1.SelectedIndex >= 1) // период с по
                {
                    if (checkBox10.Checked && !checkBox16.Checked) // разбивка продуктов по сменам
                        return 11;
                    if (checkBox16.Checked && !checkBox10.Checked) // брак
                        return 12;
                    if (checkBox16.Checked && checkBox10.Checked) // брак разбивка по сменам
                        return 13;
                    return 6;
                }

                if (comboBox1.SelectedIndex == 1) // сутки
                    return checkBox16.Checked ? 12 : 7; // брак / продукты

                if (comboBox1.SelectedIndex == 0) // сутки
                {
                    if (!checkBox16.Checked) // продукт
                    {
                        if (now.Hour >= 9 && now.Hour < 21) // за ночь
                            return 8;

                        if (now.Hour <= 9 || now.Hour >= 21) // за день
                            return 9;
                    }
                    else
                    {
                        return 12; // брак
                    }
                }
            }

            if (radioButton2.Checked)
            {
                if (!checkBox21.Checked)
                {
                    if (comboBox1.SelectedIndex >= 1) // период с по
                    {
                        if (checkBox15.Checked && !checkBox1.Checked) // разбивка простоев по типам
                            return 22;
                        else if (checkBox15.Checked && checkBox1.Checked) // разбивка простоев по типам и сменам
                            return 24;
                        else if (!checkBox15.Checked && checkBox1.Checked) // разбивка простоев по сменам
                            return 23;
                    }
                    else if (comboBox1.SelectedIndex == 0) // предыдущая смена
                    {
                        if (!checkBox15.Checked) // без разбивки простоев по типам
                        {
                            if (now.Hour >= 9 && now.Hour < 21) // за ночь
                                return 27;
                            else if (now.Hour <= 9 || now.Hour >= 21) // за день
                                return 26;
                        }
                        else // разбивка простоев по типам
                        {
                            if (now.Hour >= 9 && now.Hour < 21) // за ночь
                                return 29;
                            else if (now.Hour <= 9 || now.Hour >= 21) // за день
                                return 28;
                        }
                    }
                }
                else if (checkBox21.Checked)
                    return 25;
                // Возвращаем значение по умолчанию, если ни одно из условий не сработало
                return 21;
            }
            return 0;
        }


        private void button8_Click(object sender, EventArgs e)
        {
            (DateTime varDate1, DateTime varDate2) = GetDateRange();
            DateTime now = DateTime.Now;

            int queryNumb = GetQueryNumber(now);

            int numShift1 = checkBox9.Checked ? 1 : 0;
            int numShift2 = checkBox8.Checked ? 2 : 0;
            int numShift3 = checkBox7.Checked ? 3 : 0;
            int numShift4 = checkBox6.Checked ? 4 : 0;

            string typeStops1 = checkBox14.Checked ? "Технологические" : "_";
            string typeStops2 = checkBox13.Checked ? "Настройки" : "_";
            string typeStops3 = checkBox12.Checked ? "Поломки" : "_";
            string typeStops4 = checkBox11.Checked ? "ППР" : "_";

            Form7 form7 = new Form7(varDate1, varDate2, queryNumb, numShift1, numShift2, numShift3, numShift4,
                typeStops1, typeStops2, typeStops3, typeStops4);
            form7.Show();
        }

        private void groupBox4_Enter(object sender, EventArgs e)
        {

        }

        private void Form2_Load(object sender, EventArgs e)
        {

        }

        private void checkBox21_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void checkBox21_CheckedChanged_1(object sender, EventArgs e)
        {

        }

        private void groupBox6_Enter(object sender, EventArgs e)
        {

        }


        private void checkBox22_CheckedChanged(object sender, EventArgs e)
        {

        }
    }
}
