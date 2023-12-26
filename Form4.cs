using Microsoft.EntityFrameworkCore;
using Npgsql;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace ShiftReportApp1
{
    public partial class Form4 : Form
    {
        private System.Windows.Forms.ToolTip toolTip4;
        public Form4()
        {
            InitializeComponent();
            LoadList();
            comboBox3.Enabled = false;
            InitializeToolTip();
        }
        private void InitializeToolTip()
        {
            // Создаем экземпляр ToolTip
            toolTip4 = new System.Windows.Forms.ToolTip();

            // Настроим параметры ToolTip
            toolTip4.AutoPopDelay = 5000; // Время в миллисекундах, как долго подсказка видна после того, как мышь ушла
            toolTip4.InitialDelay = 500; // Время в миллисекундах, как долго мышь должна находиться над элементом, прежде чем появится подсказка
            toolTip4.ReshowDelay = 400; // Время в миллисекундах перед появлением подсказки, если пользователь вернулся снова

            // Задаем текст подсказки для элементов формы
            toolTip4.SetToolTip(groupBox1, "Введите дату, время суток и номер смены чтобы создать новую запись,\nесли такой смены нет в выпадающем меню слева");
            toolTip4.SetToolTip(comboBox1, "Выбирите из смен текущую, если она есть.");
            toolTip4.SetToolTip(button2, "Это кнопка отображения мини таблицы.\nВыбирите смену, которую хотите увидеть и нажмите кнопку");
            toolTip4.SetToolTip(textBox5, "Выбор места работает только в случае поломки,\nкороткой остановки, настройки");
            toolTip4.SetToolTip(checkBox1, "Если выпуск продукции не останавливался");
            toolTip4.SetToolTip(comboBox10, "Если была замена фуги, укажите номер установленной");

            // Добавьте подсказки для других элементов формы, если необходимо
        }
        // Обработчик события при наведении мыши на элемент
        private void groupBox1_MouseHover(object sender, EventArgs e) { toolTip4.Show("Привет, это подсказка!", groupBox1); }
        private void comboBox1_MouseHover(object sender, EventArgs e) { toolTip4.Show("Привет, это подсказка!", comboBox1); }
        private void button2_MouseHover(object sender, EventArgs e) { toolTip4.Show("Привет, это подсказка!", button2); }
        private void button1_MouseHover(object sender, EventArgs e) { toolTip4.Show("Привет, это подсказка!", textBox5); }
        private void checkBox1_MouseHover(object sender, EventArgs e) { toolTip4.Show("Привет, это подсказка!", checkBox1); }
        private void comboBox10_MouseHover(object sender, EventArgs e) { toolTip4.Show("Привет, это подсказка!", comboBox10); }
        protected void LoadList() // Загрузка выпадающих списков
        {
            // Блок получения даты
            DateTime now = DateTime.Now;
            DateTime varDate1 = now.AddDays(-4);
            DateTime varDate2 = now;

            // Блок получения выпадающих списков
            using (var dbContext = new ShiftReportDbContext())
            {
                // Блок выпадающего списка даты-смены
                var result1 = (
                        from sr in dbContext.ShiftReport
                        where sr.ShiftDate >= varDate1 && sr.ShiftDate <= varDate2
                        group new { sr } by new { sr.ShiftReportID, sr.ShiftDate, sr.ShiftNum, sr.ShiftDay } into grp
                        orderby grp.Key.ShiftReportID
                        select new
                        {
                            ID = grp.Key.ShiftReportID,
                            SheftDate = grp.Key.ShiftDate,
                            ShiftNum = grp.Key.ShiftNum,
                            ShiftDay = grp.Key.ShiftDay,

                        }).ToList();
                foreach (var report in result1)
                {
                    int ID = report.ID;
                    DateTime shiftDate = report.SheftDate;
                    int shiftNum = report.ShiftNum;
                    int shiftDay = report.ShiftDay;

                    string formattedDate = shiftDate.ToString("yyyy-MM-dd");

                    string outputString;
                    if (shiftDay == 1)
                    {
                        outputString = $"{ID} {formattedDate} День Смена #{shiftNum}";
                    }
                    else outputString = $"{ID} {formattedDate} Ночь Смена #{shiftNum}";
                    comboBox1.Items.Add(outputString);
                }
                // Блок выпадающего списка продуктов
                var result2 = (
                        from st in dbContext.StopType
                        group new { st } by new { st.StopName } into grp
                        select new
                        {
                            DefectCategory = grp.Key.StopName
                        }).ToList();
                foreach (var report in result2)
                {
                    string stopCat = report.DefectCategory;

                    comboBox2.Items.Add(stopCat);
                }
                // Блок выпадающих списков дефектов
                var result3 = (
                        from pil in dbContext.PlaceInLine
                        group new { pil } by new { pil.PlacesName, pil.PlacesID } into grp
                        orderby grp.Key.PlacesID
                        select new
                        {
                            PlaceName = grp.Key.PlacesName
                        }).ToList();
                foreach (var report in result3)
                {
                    string defectName = report.PlaceName;

                    comboBox3.Items.Add(defectName);
                }
                dbContext.Dispose();
            }
        }
        private void FillDropDownListsAsync()
        {
            // ... код заполнения выпадающих списков обновление
            comboBox1.Items.Clear();
            using (var dbContext = new ShiftReportDbContext())
            {
                DateTime now = DateTime.Now;
                DateTime varDate1 = now.AddDays(-4);
                DateTime varDate2 = now;
                var result = (
                        from sr in dbContext.ShiftReport
                        where sr.ShiftDate >= varDate1 && sr.ShiftDate <= varDate2
                        group new { sr } by new { sr.ShiftReportID, sr.ShiftDate, sr.ShiftNum, sr.ShiftDay } into grp
                        orderby grp.Key.ShiftReportID
                        select new
                        {
                            ID = grp.Key.ShiftReportID,
                            SheftDate = grp.Key.ShiftDate,
                            ShiftNum = grp.Key.ShiftNum,
                            ShiftDay = grp.Key.ShiftDay,

                        }).ToList();
                foreach (var report in result)
                {
                    int ID = report.ID;
                    DateTime shiftDate = report.SheftDate;
                    int shiftNum = report.ShiftNum;
                    int shiftDay = report.ShiftDay;

                    string formattedDate = shiftDate.ToString("yyyy-MM-dd");

                    string outputString;
                    if (shiftDay == 1)
                    {
                        outputString = $"{ID} {formattedDate} День Смена #{shiftNum}";
                    }
                    else outputString = $"{ID} {formattedDate} Ночь Смена #{shiftNum}";
                    comboBox1.Items.Add(outputString);
                }
                dbContext.Dispose();
            }
        }

        private void CreateDefectShiftReport() // Метод записи внесенных данных
        {
            try
            {
                ProjectLogger.LogDebug("Начало CreateShiftReport");
                using (var dbContext = new ShiftReportDbContext())
                {
                    var shiftstopreport = new StopsReport { };
                    DateTime shiftDate = (DateTime)dateTimePicker1.Value.Date;
                    int shiftDay = domainUpDown1.Text == "День" ? 1 : 2;
                    int shiftNum = (int)numericUpDown1.Value;
                    int fuge = 0;
                    if (comboBox10.Text == "1") fuge = 1;
                    else if (comboBox10.Text == "2") fuge = 2;
                    else if (comboBox10.Text == "3") fuge = 3;
                    else if (comboBox10.Text == "4") fuge = 4;
                    DateTime changeFuge = DateTime.MinValue;
                    if (comboBox10.Text != "")
                    {
                        changeFuge = new DateTime(
                            dateTimePicker1.Value.Year,
                            dateTimePicker1.Value.Month,
                            dateTimePicker1.Value.Day,
                            int.Parse(comboBox5.Text),
                            int.Parse(comboBox7.Text),
                            0  // секунды будут 0
                            );
                    }
                    int shiftID;
                    // Проверяем, существует ли запись с такими параметрами в последних 50 записях ShiftReport
                    var existingShiftReport = dbContext.ShiftReport
                        .OrderByDescending(sr => sr.ShiftReportID)
                        .Where(sr =>
                            sr.ShiftDate == shiftDate &&
                            sr.ShiftDay == shiftDay &&
                            sr.ShiftNum == shiftNum)
                        .Take(50)
                        .FirstOrDefault();

                    if (existingShiftReport != null)
                    {
                        // Запись уже существует, возвращаем её ID
                        shiftID = existingShiftReport.ShiftReportID;
                    }
                    else
                    {
                        if (string.IsNullOrWhiteSpace(comboBox1.Text))
                        {
                            // Запись не существует, создаем новую
                            var shiftReport = new ShiftReport
                            {
                                ShiftDate = shiftDate,
                                ShiftDay = shiftDay,
                                ShiftNum = shiftNum
                            };
                            dbContext.ShiftReport.Add(shiftReport);
                            dbContext.SaveChanges();
                        }
                        shiftID = (dbContext.ShiftReport.Max(sr => (int?)sr.ShiftReportID) ?? 0);
                    }
                    var defect = dbContext.StopType.FirstOrDefault(st => st.StopName == comboBox2.Text);
                    var place = dbContext.PlaceInLine.FirstOrDefault(pil => pil.PlacesName == comboBox3.Text);
                    if (defect != null)
                    {
                        // Перевод в время с форматом "HH:mm"
                        string startBreakdown = $"{comboBox4.Text}:{comboBox6.Text}";
                        DateTime startTime = DateTime.ParseExact(startBreakdown, "HH:mm", CultureInfo.InvariantCulture);
                        string endBreakdown = $"{comboBox5.Text}:{comboBox7.Text}";
                        DateTime endTime = DateTime.ParseExact(endBreakdown, "HH:mm", CultureInfo.InvariantCulture);
                        // Вычетаем и переводим в число
                        TimeSpan timeDifference = endTime.Subtract(startTime);
                        int differenceInMinutes = endTime > startTime ? (int)timeDifference.TotalMinutes : 1440 + (int)timeDifference.TotalMinutes;
                        if (!string.IsNullOrWhiteSpace(comboBox10.Text))
                        {
                            shiftstopreport = new StopsReport
                            {
                                ShiftReport = string.IsNullOrWhiteSpace(comboBox1.Text) ?
                                    shiftID : int.Parse(comboBox1.Text.Split(' ')[0]),
                                StopType = defect.StopTypeID,
                                StopFirstTime = startBreakdown,
                                StopEndTime = endBreakdown,
                                CommentStop = (string)textBox11.Text,
                                DurationStopMin = differenceInMinutes,
                                PlaceStop = comboBox3.Enabled ? place.PlacesID : 71,
                                DateRecordSR = DateTime.Now,
                                BreakdownWithoutStop = checkBox1.Checked,
                                Centrifuge = fuge
                            };
                        }
                        else
                        {
                            shiftstopreport = new StopsReport
                            {
                                ShiftReport = string.IsNullOrWhiteSpace(comboBox1.Text) ?
                                    shiftID : int.Parse(comboBox1.Text.Split(' ')[0]),
                                StopType = defect.StopTypeID,
                                StopFirstTime = startBreakdown,
                                StopEndTime = endBreakdown,
                                CommentStop = (string)textBox11.Text,
                                DurationStopMin = differenceInMinutes,
                                PlaceStop = comboBox3.Enabled ? place.PlacesID : 71,
                                DateRecordSR = DateTime.Now,
                                BreakdownWithoutStop = checkBox1.Checked
                            };
                        }

                        DialogResult result = CustomMessageBox.Show(
                            $"Дата:   {(string.IsNullOrWhiteSpace(comboBox1.Text) ? shiftDate.Date.ToString() : comboBox1.Text.Split(' ')[1])}\n" +
                            $"Номер смены:   {(string.IsNullOrWhiteSpace(comboBox1.Text) ? shiftNum.ToString() : comboBox1.Text.Split(' ')[4])}\n" +
                            $"{(string.IsNullOrWhiteSpace(comboBox1.Text) ? (shiftDay == 1 ? "День" : "Ночь") : comboBox1.Text.Split(' ')[2])}\n" +
                            $"Остановка:   {defect.StopName}\n" +
                            $"Тип остановки:   {defect.StopCategory}\n" +
                            $"Начало остановки:   {startBreakdown}\n" +
                            $"Конец остановки:   {endBreakdown}\n" +
                            $"Длительность:   {differenceInMinutes}\n" +
                            $"Комментарий:   {shiftstopreport.CommentStop}\n" +
                            $"Место отсановки:   {(comboBox3.Enabled ? place.PlacesName : "")}\n" +
                            $"Простой линии:   {(shiftstopreport.BreakdownWithoutStop ? "Нет" : "Да")}\n" +
                            $"Замена центрифуги:   {(shiftstopreport.Centrifuge > 0 ? "Да" : "Нет")}\n" +
                            $"Установлена центрифуга:   {(shiftstopreport.Centrifuge > 0 ?shiftstopreport.Centrifuge : 0)}"
                            , "Подтверждение введенных данных",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information,
                            new Font("Arial", 14));

                        if (result == DialogResult.OK)
                        {
                            dbContext.StopsReport.Add(shiftstopreport);
                            dbContext.SaveChanges();
                            MessageBox.Show("Данные успешно сохранены", "Успешно", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            SmallGrid();
                        }
                        else
                        {
                            return;
                        }
                    }
                }
                // Очистка элементов управления
                comboBox1.Text = default;
                comboBox2.Text = "";
                comboBox3.Text = default;
                domainUpDown1.Text = "День";
                checkBox1.Checked = false;
                checkBox2.Checked = false;
                comboBox4.Text = comboBox5.Text = comboBox6.Text = comboBox7.Text = textBox10.Text = textBox11.Text = comboBox10.Text = "";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Произошла ошибка: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                ProjectLogger.LogException("Ошибка в CreateShiftReport", ex);
            }
        }

        private void DifferenceMinutes()
        {
            if (comboBox4.Text != "" && comboBox5.Text != "" && comboBox6.Text != "" && comboBox7.Text != "")
            {
                string startBreakdown = $"{comboBox4.Text}:{comboBox6.Text}";
                DateTime startTime = DateTime.ParseExact(startBreakdown, "HH:mm", CultureInfo.InvariantCulture);

                string endBreakdown = $"{comboBox5.Text}:{comboBox7.Text}";
                DateTime endTime = DateTime.ParseExact(endBreakdown, "HH:mm", CultureInfo.InvariantCulture);
                // Вычетаем и переводим в число
                
                TimeSpan timeDifference = endTime.Subtract(startTime);
                int differenceInMinutes = endTime > startTime ? (int)timeDifference.TotalMinutes : 1440 + (int)timeDifference.TotalMinutes;
                textBox10.Text = $"{differenceInMinutes}";
                
            }
            else textBox10.Text = $"не указано";
        }

        private void SmallGrid()
        {
            var reportNumber = 0;

            using (var dbContext = new ShiftReportDbContext())
            {
                if (!string.IsNullOrWhiteSpace(comboBox1.Text) && int.TryParse(comboBox1.Text.Split(' ')[0], out reportNumber))
                {
                    var stopsReports = (
                        from sr in dbContext.ShiftReport
                        join str in dbContext.StopsReport on sr.ShiftReportID equals str.ShiftReport
                        join st in dbContext.StopType on str.StopType equals st.StopTypeID
                        join pil in dbContext.PlaceInLine on str.PlaceStop equals pil.PlacesID
                        where sr.ShiftReportID == reportNumber
                        group new { sr, str, st, pil } by new
                        {
                            sr.ShiftDate,
                            sr.ShiftNum,
                            str.StopReportID,
                            st.StopName,
                            pil.PlacesName,
                            str.DurationStopMin,
                            str.StopFirstTime,
                            str.StopEndTime,
                            str.CommentStop,
                            str.BreakdownWithoutStop
                        }
                        into grp
                        orderby grp.Key.StopReportID
                        select new
                        {
                            Date = grp.Key.ShiftDate,
                            ShiftNumber = grp.Key.ShiftNum,
                            StopsNumber = grp.Key.StopReportID,
                            StopName = grp.Key.StopName,
                            StopPlace = grp.Key.PlacesName,
                            StopDuration = grp.Key.DurationStopMin,
                            StopStart = grp.Key.StopFirstTime,
                            StopEndTime = grp.Key.StopEndTime,
                            Comments = grp.Key.CommentStop,
                            StopLine = grp.Key.BreakdownWithoutStop

                        }).ToList();

                    DataTable dataTable = new DataTable();
                    dataTable.Columns.Add("Дата", typeof(DateTime));
                    dataTable.Columns.Add("Смена", typeof(int));
                    dataTable.Columns.Add("#записи", typeof(int));
                    dataTable.Columns.Add("Остановка", typeof(string));
                    dataTable.Columns.Add("Место", typeof(string));
                    dataTable.Columns.Add("Длит -сть", typeof(int));
                    dataTable.Columns.Add("Начало", typeof(string));
                    dataTable.Columns.Add("Конец", typeof(string));
                    dataTable.Columns.Add("Коммент", typeof(string));
                    dataTable.Columns.Add("Был простой", typeof(bool));

                    // Заполняем DataTable
                    foreach (var report in stopsReports)
                    {
                        DataRow row = dataTable.NewRow();
                        row["Дата"] = report.Date;
                        row["смена"] = report.ShiftNumber;
                        row["#записи"] = report.StopsNumber;
                        row["Остановка"] = report.StopName;
                        row["Место"] = report.StopPlace;
                        row["Длит -сть"] = report.StopDuration;
                        row["Начало"] = report.StopStart;
                        row["Конец"] = report.StopEndTime;
                        row["Коммент"] = report.Comments;
                        row["Был простой"] = report.StopLine;
                        // Заполняйте остальные поля аналогично
                        dataTable.Rows.Add(row);

                        for (int i = 0; i> dataTable.Rows.Count; i++ )
                        {
                            dataTable.Rows[i][6] = dataTable.Rows[i][6].ToString().Split(' ')[1];
                            dataTable.Rows[i][7] = dataTable.Rows[i][7].ToString().Split(' ')[1];
                        }

                        dataGridView1.DataSource = dataTable;
                    }
                }
                else
                {
                    dataGridView1.Text = "Неверный формат отчета.";
                }

                dataGridView1.ColumnHeadersDefaultCellStyle.Font = new Font("Arial", 7); // Устанавливаем шрифт и размер
                dataGridView1.DefaultCellStyle.Font = new Font("Arial", 8); // Устанавливаем шрифт и размер
                dataGridView1.RowTemplate.Height = 16; // Устанавливаем высоту строки
                dataGridView1.ColumnHeadersHeight = 16; // Устанавливаем высоту заголовка столбца
                dataGridView1.CellBorderStyle = DataGridViewCellBorderStyle.Single;
                if (dataGridView1.Rows.Count > 0)
                {
                    dataGridView1.Columns[0].Width = 65;
                    dataGridView1.Columns[1].Width = 35;
                    dataGridView1.Columns[2].Width = 35;
                    dataGridView1.Columns[3].Width = 65;
                    dataGridView1.Columns[4].Width = 50;
                    dataGridView1.Columns[5].Width = 40;
                    dataGridView1.Columns[6].Width = 50;
                    dataGridView1.Columns[7].Width = 50;
                    dataGridView1.Columns[8].Width = 50;
                    dataGridView1.Columns[9].Width = 40;
                }
                dbContext.Dispose();
            }
        }

        private System.Windows.Forms.ToolTip cellToolTip = new System.Windows.Forms.ToolTip();

        private void dataGridView1_CellMouseEnter(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0 && e.ColumnIndex >= 0)
            {
                DataGridViewCell cell = dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex];
                string cellContent = cell.Value != null ? cell.Value.ToString() : string.Empty;

                // Показываем всплывающее окно только если содержимое ячейки не пустое
                if (!string.IsNullOrEmpty(cellContent))
                {
                    cellToolTip.Show(cellContent, dataGridView1, cell.ContentBounds.Right, cell.ContentBounds.Bottom, 1000);
                }
            }
        }

        private void dataGridView1_CellMouseLeave(object sender, DataGridViewCellEventArgs e)
        {
            cellToolTip.Hide(dataGridView1);
        }
        private void button5_Click(object sender, EventArgs e)
        {
            DateTime now = DateTime.Now;
            DateTime getDate1 = default;
            DateTime getDate2 = default;
            List<int> shiftsDay = new List<int> { };
            List<int> shifts = new List<int> { };
            List<string> stopCategoryes = new List<string> { };
            List<string> prodList = new List<string> { };

            if (now.Hour >= 0 && now.Hour <= 9) { getDate1 = now.AddDays(-1).Date; getDate2 = now.AddDays(-1).Date; shiftsDay.Add(2); }
            else if (now.Hour < 21 && now.Hour > 9) { getDate1 = now.Date; getDate2 = now.Date; shiftsDay.Add(1); }
            if (now.Hour >= 21 && now.Hour < 24) { getDate1 = now.Date; getDate2 = now.Date; shiftsDay.Add(2); }
            int getMethod = 15;
            SaveInXML saver = new SaveInXML();  // Создаем экземпляр класса SaAsDi
            saver.SaveExcelFile(getMethod, prodList, shiftsDay, shifts, stopCategoryes, getDate1, getDate2); // Вызываем метод сохранения
        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void textBox7_TextChanged(object sender, EventArgs e)
        {

        }

        private void dateTimePicker3_ValueChanged(object sender, EventArgs e)
        {

        }

        private void textBox12_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox11_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox14_TextChanged(object sender, EventArgs e)
        {

        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {

        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {

        }

        private void domainUpDown1_SelectedItemChanged(object sender, EventArgs e)
        {

        }

        private void numericUpDown1_ValueChanged(object sender, EventArgs e)
        {

        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox2.SelectedItem.Equals("Поломка") || comboBox2.SelectedItem.Equals("Настройка оборудования") || comboBox2.SelectedItem.Equals("Короткая остановка до 5 минут"))
            {
                comboBox3.Enabled = true;
            }
            else comboBox3.Enabled = false;
        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {

        }

        private void groupBox2_Enter(object sender, EventArgs e)
        {

        }

        private void textBox5_TextChanged(object sender, EventArgs e)
        {

        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void textBox6_TextChanged(object sender, EventArgs e)
        {

        }

        private void dateTimePicker2_ValueChanged(object sender, EventArgs e)
        {

        }

        private void textBox9_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox10_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox13_TextChanged(object sender, EventArgs e)
        {

        }

        private void button6_Click(object sender, EventArgs e)
        {
            CreateDefectShiftReport();
            FillDropDownListsAsync();
            SmallGrid();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            Form4 form4 = Application.OpenForms.OfType<Form4>().FirstOrDefault();
            form4.Close();
        }

        private void Form4_Load(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            SmallGrid();
        }

        private void groupBox3_Enter(object sender, EventArgs e)
        {

        }

        private void groupBox4_Enter(object sender, EventArgs e)
        {

        }

        private void textBox13_TextChanged_1(object sender, EventArgs e)
        {

        }

        private void textBox15_TextChanged(object sender, EventArgs e)
        {

        }

        private void comboBox9_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void comboBox8_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void comboBox10_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void comboBox5_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            DifferenceMinutes();
        }

        private void comboBox7_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            DifferenceMinutes();
        }

        private void comboBox4_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            DifferenceMinutes();
        }

        private void comboBox6_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            DifferenceMinutes();
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            DateTime now = DateTime.Now;
            DateTime getDate1 = default;
            DateTime getDate2 = default;
            List<int> shiftsDay = new List<int> { };
            List<int> shifts = new List<int> { };
            List<string> stopCategoryes = new List<string> { };
            List<string> prodList = new List<string> { };

            if (now.Hour >= 0 && now.Hour <= 9) { getDate1 = now.AddDays(-1).Date; getDate2 = now.AddDays(-1).Date; shiftsDay.Add(2); }
            else if (now.Hour < 21 && now.Hour > 9) { getDate1 = now.Date; getDate2 = now.Date; shiftsDay.Add(1); }
            if (now.Hour >= 21 && now.Hour < 24) { getDate1 = now.Date; getDate2 = now.Date; shiftsDay.Add(2); }
            int getMethod = 21;

            SaveInXML saver = new SaveInXML();  // Создаем экземпляр класса SaAsDi
            saver.SaveExcelFile(getMethod, prodList, shiftsDay, shifts, stopCategoryes, getDate1, getDate2); // Вызываем метод сохранения
        }
    }
}
