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
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace ShiftReportApp1
{
    public partial class Form2 : Form
    {
        private System.Windows.Forms.ToolTip toolTip2;
        public Form2()
        {
            InitializeComponent();
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

            checkedListBox1.Enabled = false;
            checkedListBox2.Enabled = false;
            button1.Enabled = false;
            button2.Enabled = false;
            button3.Enabled = false;
            button4.Enabled = false;

            checkBox6.Enabled = false;
            checkBox7.Enabled = false;
            checkBox8.Enabled = false;
            checkBox9.Enabled = false;
            checkBox20.Enabled = false;
            InitializeToolTip();
    }
        private void InitializeToolTip()
        {
            // Создаем экземпляр ToolTip
            toolTip2 = new System.Windows.Forms.ToolTip();

            // Настроим параметры ToolTip
            toolTip2.AutoPopDelay = 5000; // Время в миллисекундах, как долго подсказка видна после того, как мышь ушла
            toolTip2.InitialDelay = 1000; // Время в миллисекундах, как долго мышь должна находиться над элементом, прежде чем появится подсказка
            toolTip2.ReshowDelay = 500; // Время в миллисекундах перед появлением подсказки, если пользователь вернулся снова

            // Задаем текст подсказки для элементов формы
            toolTip2.SetToolTip(numericUpDown1, "Номер недели в году");
            toolTip2.SetToolTip(checkedListBox1, "Список продуктов, производившихся за выбранный период");
            toolTip2.SetToolTip(checkedListBox2, "Список продуктов, по которым нужно отобразить статистику");
            toolTip2.SetToolTip(button1, "Переместить весь список продуктов");
            toolTip2.SetToolTip(button3, "Переместить весь список продуктов");
            toolTip2.SetToolTip(button2, "Переместить выбранные продукты");
            toolTip2.SetToolTip(button4, "Переместить выбранные продукты");

        }
        // Обработчик события при наведении мыши на элемент
        private void numericUpDown1_MouseHover(object sender, EventArgs e) { toolTip2.Show("Привет, это подсказка!", numericUpDown1); }
        private void checkedListBox1_MouseHover(object sender, EventArgs e) { toolTip2.Show("Привет, это подсказка!", checkedListBox1); }
        private void checkedListBox2_MouseHover(object sender, EventArgs e) { toolTip2.Show("Привет, это подсказка!", checkedListBox2); }
        private void button1_MouseHover(object sender, EventArgs e) { toolTip2.Show("Привет, это подсказка!", button1); }
        private void button3_MouseHover(object sender, EventArgs e) { toolTip2.Show("Привет, это подсказка!", button3); }
        private void button2_MouseHover(object sender, EventArgs e) { toolTip2.Show("Привет, это подсказка!", button2); }
        private void button4_MouseHover(object sender, EventArgs e) { toolTip2.Show("Привет, это подсказка!", button4); }

        private void ExtractProduct(DateTime varDate1, DateTime varDate2, IEnumerable<int> shiftDays)
        {
            try
            {
                ProjectLogger.LogDebug("Начало ExtractProduct");
                checkedListBox1.Items.Clear();
                checkedListBox2.Items.Clear();

                LINQRequest lINQRequest = new LINQRequest();
                List<ItemReportList> prodQualityReports = lINQRequest.SetReportList(varDate1, varDate2, shiftDays);

                    foreach (var report in prodQualityReports)
                    {
                        string productName = report.ProductNames;
                        bool unspecified = report.Unspecifies;
                        int prodDepth = report.Depth;
                        int length = report.Length;
                        int width = report.Width;

                        string outputString;

                        if (productName.Contains("М"))
                        {
                            if (unspecified)
                            {

                                string lastThreeChars = productName.Split('М')[1];
                                outputString = $"{productName.Split('М')[0]} {length}x{width}x{prodDepth} ({lastThreeChars})";
                            }
                            else
                            {
                            outputString = $"{productName} {length}x{width}x{prodDepth}";
                            }
                        }
                        else
                        {
                            outputString = $"{productName} {length}x{width}x{prodDepth}";
                        }

                        checkedListBox1.Items.Add(outputString);
                    }
               
                ProjectLogger.LogDebug("Конец ExtractProduct");
            }
            catch (Exception ex)
            {
                // Обработка ошибки
                MessageBox.Show($"An error occurred: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                ProjectLogger.LogException("Ошибка в ExtractProduct", ex);
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
                        varDate2 = now.Date.AddHours(8);
                    }
                    if (now.Hour >= 21 && now.Hour <= 23)
                    {
                        varDate1 = now.Date.AddHours(8);
                        varDate2 = now.Date.AddDays(1).AddHours(8);
                    }
                    if (0 <= now.Hour && now.Hour < 9)
                    {
                        varDate1 = now.Date.AddDays(-1).AddHours(8);
                        varDate2 = now.Date.AddHours(8);
                    }
                }
                else if (comboBox1.SelectedIndex == 1)
                {
                    varDate1 = now.Date.AddDays(-1);
                    varDate2 = now.Date.AddDays(-1);
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
            checkedListBox1.Enabled = false;
            checkedListBox2.Enabled = false;
            button1.Enabled = false;
            button2.Enabled = false;
            button3.Enabled = false;
            button4.Enabled = false;
            checkedListBox1.Items.Clear();
            checkedListBox2.Items.Clear();

            // Проверяем выбранный элемент в ComboBox
            switch (comboBox1.SelectedIndex)
            {
                case 2: // С даты по дату
                    dateTimePicker1.Enabled = true;
                    dateTimePicker2.Enabled = true;
                    groupBox4.Enabled = true;
                    groupBox3.Enabled = true;
                    checkedListBox1.Enabled = true;
                    checkedListBox2.Enabled = true;
                    button1.Enabled = true;
                    button2.Enabled = true;
                    button3.Enabled = true;
                    button4.Enabled = true;
                    break;
                case 3: // За неделю
                    numericUpDown1.Enabled = true;
                    groupBox4.Enabled = true;
                    groupBox3.Enabled = true;
                    checkedListBox1.Enabled = true;
                    checkedListBox2.Enabled = true;
                    button1.Enabled = true;
                    button2.Enabled = true;
                    button3.Enabled = true;
                    button4.Enabled = true;
                    break;
                case 5: // За месяц
                    numericUpDown3.Enabled = true;
                    groupBox4.Enabled = true;
                    groupBox3.Enabled = true;
                    checkedListBox1.Enabled = true;
                    checkedListBox2.Enabled = true;
                    button1.Enabled = true;
                    button2.Enabled = true;
                    button3.Enabled = true;
                    button4.Enabled = true;
                    break;
                case 4: // За год
                    numericUpDown2.Enabled = true;
                    groupBox4.Enabled = true;
                    groupBox3.Enabled = true;
                    checkedListBox1.Enabled = true;
                    checkedListBox2.Enabled = true;
                    button1.Enabled = true;
                    button2.Enabled = true;
                    button3.Enabled = true;
                    button4.Enabled = true;
                    break;
                default: // За предыдущую смену и За предыдущие сутки
                    break;
            }

            DateTime now = DateTime.Now;

            if (radioButton1.Checked && now.Hour >= 9 && now.Hour < 21 && comboBox1.SelectedIndex == 0)
            {
                (DateTime varDate1, DateTime varDate2) = GetDateRange();
                List<int> shiftDays = new List<int> { 2 };
                ExtractProduct(varDate1, varDate2, shiftDays); // за предыдущую смену после 9-00
            }
            else if (radioButton1.Checked && now.Hour <= 9 && comboBox1.SelectedIndex == 0)
            {
                (DateTime varDate1, DateTime varDate2) = GetDateRange();
                List<int> shiftDays = new List<int> { 1 };
                ExtractProduct(varDate1, varDate2, shiftDays); // за предыдущую смену после 00-00 до 9-00
            }
            else if (radioButton1.Checked && now.Hour >= 21 && comboBox1.SelectedIndex == 0)
            {
                (DateTime varDate1, DateTime varDate2) = GetDateRange();
                List<int> shiftDays = new List<int> { 1 };
                ExtractProduct(varDate1, varDate2, shiftDays); //за предыдущую смену после 21-00 до 23-59
            }
            else if (radioButton1.Checked && comboBox1.SelectedIndex == 1)
            {
                (DateTime varDate1, DateTime varDate2) = GetDateRange();
                List<int> shiftDays = new List<int> { 1, 2 };
                ExtractProduct(varDate1, varDate2, shiftDays); // за предыдущие сутки после 9-00 до 23-59
            }
            else if (radioButton1.Checked && comboBox1.SelectedIndex == 2)
            {
                (DateTime varDate1, DateTime varDate2) = GetDateRange();
                List<int> shiftDays = new List<int> { 1, 2 };
                ExtractProduct(varDate1, varDate2, shiftDays);  // с даты по дату
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
                (DateTime varDate1, DateTime varDate2) = GetDateRange();
                List<int> shiftDays = new List<int> { 1, 2 };
                ExtractProduct(varDate1, varDate2, shiftDays);
            }
        }

        private void dateTimePicker2_ValueChanged(object sender, EventArgs e)
        {
            if (radioButton1.Checked)
            {
                (DateTime varDate1, DateTime varDate2) = GetDateRange();
                List<int> shiftDays = new List<int> { 1, 2 };
                ExtractProduct(varDate1, varDate2, shiftDays);
            }
        }

        private void numericUpDown1_ValueChanged(object sender, EventArgs e)
        {
            if (radioButton1.Checked)
            {
                (DateTime varDate1, DateTime varDate2) = GetDateRange();
                List<int> shiftDays = new List<int> { 1, 2 };
                ExtractProduct(varDate1, varDate2, shiftDays);
            }
        }

        private void numericUpDown2_ValueChanged(object sender, EventArgs e)
        {
            if (radioButton1.Checked)
            {
                (DateTime varDate1, DateTime varDate2) = GetDateRange();
                List<int> shiftDays = new List<int> { 1, 2 };
                ExtractProduct(varDate1, varDate2, shiftDays);
            }
        }

        private void numericUpDown3_ValueChanged(object sender, EventArgs e)
        {
            if (radioButton1.Checked)
            {
                (DateTime varDate1, DateTime varDate2) = GetDateRange();
                List<int> shiftDays = new List<int> { 1, 2 };
                ExtractProduct(varDate1, varDate2, shiftDays);
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
            
        }

        private void checkBox17_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void checkBox18_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void checkBox10_CheckedChanged(object sender, EventArgs e)
        {
            checkBox6.Enabled = checkBox10.Checked ? true : false;
            checkBox7.Enabled = checkBox10.Checked ? true : false;
            checkBox8.Enabled = checkBox10.Checked ? true : false;
            checkBox9.Enabled = checkBox10.Checked ? true : false;
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
            List<int> shifts = new List<int> { };
            //инициализируем состояние кнопок смен
            if (checkBox10.Checked)
            {
                int numShift1 = checkBox9.Checked ? 1 : 0;
                int numShift2 = checkBox8.Checked ? 2 : 0;
                int numShift3 = checkBox7.Checked ? 3 : 0;
                int numShift4 = checkBox6.Checked ? 4 : 0;
                shifts = new List<int> { numShift1, numShift2, numShift3, numShift4 };
            }
            else if(checkBox1.Checked)
            {
                int numShift1 = checkBox2.Checked ? 1 : 0;
                int numShift2 = checkBox3.Checked ? 2 : 0;
                int numShift3 = checkBox4.Checked ? 3 : 0;
                int numShift4 = checkBox5.Checked ? 4 : 0;
                shifts = new List<int> { numShift1, numShift2, numShift3, numShift4 };
            }

            string typeStops1 = checkBox14.Checked ? "Технологический" : "_";
            string typeStops2 = checkBox13.Checked ? "Настройки" : "_";
            string typeStops3 = checkBox12.Checked ? "Неплановый" : "_";
            string typeStops4 = checkBox11.Checked ? "Плановый" : "_";
            List<string> stopCategoryes = new List<string> { typeStops1, typeStops2, typeStops3, typeStops4 };

            // инициализируем даты
            (DateTime varDate1, DateTime varDate2) = GetDateRange();
            DateTime now = DateTime.Now;

            var resultCall = GetQueryNumber(now);
            int getMethod = resultCall.Item1;
            List<int> shiftsDay = resultCall.Item2;

            // запишем в список все отмеченные элементы в checkedListBox2
            List<string> prodList = checkedListBox2.CheckedItems.Cast<string>().ToList();

            SaveInXML saver = new SaveInXML();  // Создаем экземпляр класса SaAsDi
            saver.SaveExcelFile(getMethod, prodList, shiftsDay, shifts, stopCategoryes, varDate1, varDate2); // Вызываем метод сохранения
        }

        private void button5_Click(object sender, EventArgs e)
        {

        }

        private void button7_Click(object sender, EventArgs e)
        {
            Form2 form2 = Application.OpenForms.OfType<Form2>().FirstOrDefault();
            form2.Close();
        }

        private (int, List<int>) GetQueryNumber(DateTime now)
        {
            List<int> list = new List<int>();
            // продукты
            if (radioButton1.Checked)
            {
                if (!checkBox21.Checked)
                {
                    if (comboBox1.SelectedIndex > 1) // период с по
                    {
                        list = new List<int> { 1, 2 };
                        if (checkBox10.Checked && !checkBox16.Checked) // разбивка продуктов по сменам
                        {
                            return (3, list);
                        }
                        if (checkBox16.Checked && !checkBox10.Checked) // брак
                        {
                            return (4, list);
                        }
                        if (checkBox16.Checked && checkBox10.Checked) // брак разбивка по сменам
                        {
                            return (5, list);
                        }
                        return (0, list);
                    }

                    if (comboBox1.SelectedIndex == 1) // сутки
                    {
                        list = new List<int> { 1, 2 };
                        return (checkBox16.Checked ? 4 : 2, list); // брак / продукты
                    }
                    if (comboBox1.SelectedIndex == 0) // смену
                    {
                        if (!checkBox16.Checked) // продукт
                        {
                            if (now.Hour >= 9 && now.Hour < 21) // за ночь
                            {
                                list = new List<int> { 2 };
                                return (2, list);
                            }
                            if (now.Hour <= 9 || now.Hour >= 21) // за день
                            {
                                list = new List<int> { 1 };
                                return (2, list);
                            }
                        }
                        else
                        {
                            list = new List<int> { 1, 2 };
                            return (4, list); // брак
                        }
                    }
                }
            }
            // простои
            if (radioButton2.Checked)
            {
                if (!checkBox21.Checked)
                {
                    if (comboBox1.SelectedIndex >= 2) // период с по
                    {
                        list = new List<int> { 1, 2 };
                        if (checkBox15.Checked && !checkBox1.Checked) // разбивка простоев по типам
                        {
                            return (11, list);
                        }
                        else if (!checkBox15.Checked && checkBox1.Checked) // разбивка простоев по типам и сменам
                        {
                            return (12, list);
                        }
                        else if (checkBox15.Checked && checkBox1.Checked) // разбивка простоев по типам и сменам
                        {
                            return (14, list);
                        }
                        return (10, list);

                    }
                    else if (comboBox1.SelectedIndex == 1) // предыдущие сутки
                    {
                        list = new List<int> { 1, 2 };
                        return (15, list);
                    }
                    else if (comboBox1.SelectedIndex == 0) // предыдущая смена
                    {
                        if (!checkBox15.Checked) // без разбивки простоев по типам
                        {
                            if (now.Hour >= 9 && now.Hour < 21) // за ночь
                            {
                                list = new List<int> { 2 };
                            }
                            else if (now.Hour <= 9 || now.Hour >= 21) // за день
                            {
                                list = new List<int> { 1 };
                            }
                            return (15, list);
                        }
                        else // разбивка простоев по типам
                        {
                            if (now.Hour >= 9 && now.Hour < 21) // за ночь
                            {
                                list = new List<int> { 2 };
                            }
                            else if (now.Hour <= 9 || now.Hour >= 21) // за день
                            {
                                list = new List<int> { 1 };
                            }
                            return (14, list);
                        }
                    }
                }
            }
            if (checkBox21.Checked)
            {
                list = new List<int> { 1, 2 };
                return (13, list);
            }
            // Возвращаем значение по умолчанию, если ни одно из условий не сработало
            return (0, new List<int>());
        }


        private void button8_Click(object sender, EventArgs e)
        {
            (DateTime varDate1, DateTime varDate2) = GetDateRange();
            DateTime now = DateTime.Now;

            var resultCall = GetQueryNumber(now);
            int getMethod = resultCall.Item1;
            List<int> shiftsDay = resultCall.Item2;

            List<int> shifts = new List<int> { };
            //инициализируем состояние кнопок смен
            if (checkBox10.Checked)
            {
                int numShift1 = checkBox9.Checked ? 1 : 0;
                int numShift2 = checkBox8.Checked ? 2 : 0;
                int numShift3 = checkBox7.Checked ? 3 : 0;
                int numShift4 = checkBox6.Checked ? 4 : 0;
                shifts = new List<int> { numShift1, numShift2, numShift3, numShift4 };
            }
            else if (checkBox1.Checked)
            {
                int numShift1 = checkBox2.Checked ? 1 : 0;
                int numShift2 = checkBox3.Checked ? 2 : 0;
                int numShift3 = checkBox4.Checked ? 3 : 0;
                int numShift4 = checkBox5.Checked ? 4 : 0;
                shifts = new List<int> { numShift1, numShift2, numShift3, numShift4 };
            }

            string typeStops1 = checkBox14.Checked ? "Технологический" : "_";
            string typeStops2 = checkBox13.Checked ? "Настройки" : "_";
            string typeStops3 = checkBox12.Checked ? "Неплановый" : "_";
            string typeStops4 = checkBox11.Checked ? "Плановый" : "_";
            List<string> stopCategoryes = new List<string> { typeStops1, typeStops2, typeStops3, typeStops4 };

            Form7 form7 = new Form7(getMethod, varDate1, varDate2, shiftsDay, shifts, stopCategoryes);
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
