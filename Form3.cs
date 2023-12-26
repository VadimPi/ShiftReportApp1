using DocumentFormat.OpenXml.Office.Word;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Vml;
using Microsoft.EntityFrameworkCore;
using Npgsql;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.StartPanel;
using ComboBox = System.Windows.Forms.ComboBox;
using Font = System.Drawing.Font;

namespace ShiftReportApp1
{
    public partial class Form3 : Form
    {
        private System.Windows.Forms.ToolTip toolTip3;

        public Form3()
        {
            InitializeComponent();
            OnLoad();
            numericUpDown17.Enabled = false;
            numericUpDown8.Enabled = false;
            numericUpDown9.Enabled = false;
            numericUpDown10.Enabled = false;
            numericUpDown11.Enabled = false;
            numericUpDown12.Enabled = false;
            numericUpDown13.Enabled = false;
            numericUpDown14.Enabled = false;

            dateTimePicker1.MinDate = DateTime.Today.AddDays(-7);
            dateTimePicker1.MaxDate = DateTime.Today.AddDays(0);
            InitializeToolTip();
        }

        private void InitializeToolTip()
        {
            // Создаем экземпляр ToolTip
            toolTip3 = new System.Windows.Forms.ToolTip();

            // Настроим параметры ToolTip
            toolTip3.AutoPopDelay = 5000; // Время в миллисекундах, как долго подсказка видна после того, как мышь ушла
            toolTip3.InitialDelay = 1000; // Время в миллисекундах, как долго мышь должна находиться над элементом, прежде чем появится подсказка
            toolTip3.ReshowDelay = 500; // Время в миллисекундах перед появлением подсказки, если пользователь вернулся снова

            // Задаем текст подсказки для элементов формы
            toolTip3.SetToolTip(groupBox1, "Введите дату, время суток и номер смены чтобы создать новую запись,\nесли такой смены нет в выпадающем меню слева");
            toolTip3.SetToolTip(comboBox1, "Выбирите из смен текущую, если она есть.");
            toolTip3.SetToolTip(button2, "Это кнопка отображения мини таблицы.\nВыбирите смену, которую хотите увидеть и нажмите кнопку");
            toolTip3.SetToolTip(button1, "Это кнопка отображения расширеной таблицы в новом окне.\nВыбирите смену, которую хотите увидеть и нажмите кнопку");

        }
        // Обработчик события при наведении мыши на элемент
        private void groupBox1_MouseHover(object sender, EventArgs e) { toolTip3.Show("Привет, это подсказка!", groupBox1); }
        private void comboBox1_MouseHover(object sender, EventArgs e) { toolTip3.Show("Привет, это подсказка!", comboBox1); }
        private void button2_MouseHover(object sender, EventArgs e) { toolTip3.Show("Привет, это подсказка!", button2); }
        private void button1_MouseHover(object sender, EventArgs e) { toolTip3.Show("Привет, это подсказка!", button1); }

        public class ComboBoxNumericUpDownPair
        {
            public System.Windows.Forms.ComboBox ComboBox { get; set; }
            public NumericUpDown NumericUpDown { get; set; }
        }

        protected void OnLoad() // Загрузка выпадающих списков
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
                        from pc in dbContext.ProductCategories
                        group new { pc } by new { pc.ProductName, pc.ProductID } into grp
                        orderby grp.Key.ProductID
                        select new
                        {
                            ProductCategory = grp.Key.ProductName
                        }).ToList();
                foreach (var report in result2)
                {
                    string prodCat = report.ProductCategory;

                    comboBox2.Items.Add(prodCat);
                }
                // Блок выпадающих списков дефектов
                var result3 = (
                        from dt in dbContext.DefectTypes
                        group new { dt } by new { dt.DefectName } into grp
                        where grp.Key.DefectName != "Обрезь"
                        select new
                        {
                            DefectName = grp.Key.DefectName
                        }).ToList();
                foreach (var report in result3)
                {
                    string defectName = report.DefectName;

                    comboBox3.Items.Add(defectName);
                    comboBox4.Items.Add(defectName);
                    comboBox5.Items.Add(defectName);
                    comboBox6.Items.Add(defectName);
                    comboBox7.Items.Add(defectName);
                    comboBox8.Items.Add(defectName);
                    comboBox9.Items.Add(defectName);
                }
                dbContext.Dispose();
            }
        }
        private void FillDropDownLists()
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

        private void SmallGrid()
        {
            var reportNumber = 0;

            using (var dbContext = new ShiftReportDbContext())
            {
                if (!string.IsNullOrWhiteSpace(comboBox1.Text) && int.TryParse(comboBox1.Text.Split(' ')[0], out reportNumber))
                {
                    var prodQualityReports = dbContext.ProdQualityReport
                    .Where(pqr =>
                    dbContext.ShiftReport
                    .Where(sr => sr.ShiftReportID == reportNumber)
                    .Select(sr => sr.ShiftReportID)
                    .Contains(pqr.Report))
                    .Join(
                        dbContext.ProductCategories,
                        pqr => pqr.Product,
                        pc => pc.ProductID,
                        (pqr, pc) => new
                        {
                            ProductName = pc.ProductName,
                            Unspecified = pqr.Unspecified,
                            ProdDepth = pqr.ProdDepth,
                            Length = pqr.ProdLength,
                            Width = pqr.ProdWidth,
                            Regarding = pqr.Regarding,
                            AvgDensity = pqr.AvgDensity,
                            PackCount = pqr.PackCount,
                            Volume = Math.Round(pqr.VolumeProduct, 3),
                            Weight = Math.Round(pqr.Weight)

                        })
                    .Distinct()
                    .ToList();

                    DataTable dataTable = new DataTable();
                    dataTable.Columns.Add("Марка", typeof(string));
                    dataTable.Columns.Add("Толщ.", typeof(int));
                    dataTable.Columns.Add("Длин.", typeof(int));
                    dataTable.Columns.Add("Шир.", typeof(int));
                    dataTable.Columns.Add("Неуказ.", typeof(bool));
                    dataTable.Columns.Add("Плот-ть", typeof(float));
                    dataTable.Columns.Add("Кол-во пачек", typeof(int));
                    dataTable.Columns.Add("Объем", typeof(float));
                    dataTable.Columns.Add("Вес", typeof(float));

                    // Заполняем DataTable
                    foreach (var report in prodQualityReports)
                    {
                        DataRow row = dataTable.NewRow();
                        row["Марка"] = report.ProductName;
                        row["Толщ."] = report.ProdDepth;
                        row["Длин."] = report.Length;
                        row["Шир."] = report.Width;
                        row["Неуказ."] = report.Unspecified;
                        row["Плот-ть"] = report.AvgDensity;
                        row["Кол-во пачек"] = report.PackCount;
                        row["Объем"] = report.Volume;
                        row["Вес"] = report.Weight;
                        // Заполняйте остальные поля аналогично
                        dataTable.Rows.Add(row);

                        dataGridView1.DataSource = dataTable;
                    }
                }
                else
                {
                    dataGridView1.Text = "Неверный формат отчета.";
                }

                dataGridView1.ColumnHeadersDefaultCellStyle.Font = new Font("Arial", 6); // Устанавливаем шрифт и размер
                dataGridView1.DefaultCellStyle.Font = new Font("Arial", 6); // Устанавливаем шрифт и размер
                dataGridView1.RowTemplate.Height = 12; // Устанавливаем высоту строки
                dataGridView1.ColumnHeadersHeight = 12; // Устанавливаем высоту заголовка столбца
                dataGridView1.CellBorderStyle = DataGridViewCellBorderStyle.Single;
                if (dataGridView1.Rows.Count > 0)
                {
                    dataGridView1.Columns[0].Width = 45;
                    dataGridView1.Columns[1].Width = 20;
                    dataGridView1.Columns[2].Width = 30;
                    dataGridView1.Columns[3].Width = 30;
                    dataGridView1.Columns[4].Width = 30;
                    dataGridView1.Columns[5].Width = 30;
                    dataGridView1.Columns[6].Width = 30;
                    dataGridView1.Columns[7].Width = 35;
                    dataGridView1.Columns[8].Width = 30;
                }
                dbContext.Dispose();
            }
        }

        private void CreateProductShiftReport() // Метод записи внесенных данных
        {  
            try
            {
                ProjectLogger.LogDebug("Начало CreateShiftReport");

                using (var dbContext = new ShiftReportDbContext())
                {
                    DateTime shiftDate = (DateTime)dateTimePicker1.Value.Date;
                    int shiftDay = domainUpDown1.Text == "День" ? 1 : 2;
                    int shiftNum = (int)numericUpDown1.Value;
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
                    var product = dbContext.ProductCategories.FirstOrDefault(pc => pc.ProductName == comboBox2.Text);
                    if (product != null)
                    {
                        double volumeProd = (float)numericUpDown5.Value * (int)numericUpDown6.Value;
                        volumeProd = Math.Round(volumeProd, 3);
                        double weight = (float)numericUpDown5.Value * (int)numericUpDown6.Value * (float)numericUpDown7.Value;
                        weight = Math.Round(weight);

                        var shiftqualreport = new ProdQualityReport
                        {
                            Report = string.IsNullOrWhiteSpace(comboBox1.Text) ?
                                shiftID :
                                int.Parse(comboBox1.Text.Split(' ')[0]),
                            Product = product.ProductID,
                            ProdDepth = (int)numericUpDown2.Value,
                            ProdLength = (int)numericUpDown4.Value,
                            ProdWidth = (int)numericUpDown3.Value,
                            VolumePack = (float)numericUpDown5.Value,
                            AvgDensity = (float)numericUpDown7.Value,
                            PackCount = (int)numericUpDown6.Value,
                            Unspecified = checkBox1.Checked,
                            Regarding = checkBox2.Checked,
                            DateRecordPQR = DateTime.Now,
                            VolumeProduct = (float)volumeProd,
                            Weight = (float)weight
                        };

                        DialogResult result = CustomMessageBox.Show(
                            $"Дата:   {(string.IsNullOrWhiteSpace(comboBox1.Text) ? shiftDate.Date.ToString() : comboBox1.Text.Split(' ')[1])}\n" +
                            $"Номер смены:   {(string.IsNullOrWhiteSpace(comboBox1.Text) ? shiftNum.ToString() : comboBox1.Text.Split(' ')[4])}\n" +
                            $"{(string.IsNullOrWhiteSpace(comboBox1.Text) ? (shiftDay == 1 ? "День" : "Ночь") : comboBox1.Text.Split(' ')[2])}\n" +
                            $"Марка:   {product.ProductName}\n" +
                            $"Неуказанная:   {(shiftqualreport.Unspecified ? "Да" : "Нет")}\n" +
                            $"Пересорт:   {(shiftqualreport.Regarding ? "Да" : "Нет")}\n" +
                            $"Длинна:   {shiftqualreport.ProdLength}\n" +
                            $"Ширина:   {shiftqualreport.ProdWidth}\n" +
                            $"Толщина:   {shiftqualreport.ProdDepth}\n" +
                            $"Объем 1 пачки:   {shiftqualreport.VolumePack}\n" +
                            $"Средняя плотность:   {shiftqualreport.AvgDensity}\n" +
                            $"Количество пачек:   {shiftqualreport.PackCount}\n" +
                            $"Объем (м3):   {shiftqualreport.VolumeProduct}\n" +
                            $"Вес (кг):   {shiftqualreport.Weight}"
                            , "Подтверждение введенных данных",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information,
                            new Font("Arial", 14));

                        if (result == DialogResult.OK)
                        {
                            dbContext.ProdQualityReport.Add(shiftqualreport);
                            dbContext.SaveChanges();
                            MessageBox.Show("Данные успешно сохранены", "Успешно", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            SmallGrid();
                        }
                        else
                        {
                            return;
                        }
                    }
                    else
                    {
                        MessageBox.Show("Проверьте введенный продукт", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }

                    if (numericUpDown16.Value != 0 && numericUpDown15.Value != 0)
                    {
                        var lastPQReports = dbContext.ProdQualityReport.OrderByDescending(pq => pq.PQReportID).FirstOrDefault();
                        var prodDefectReport = new ProdDefectReport
                        {
                            ProductReport = lastPQReports.PQReportID,
                            DefectType = 8,
                            DefectVolume = (float)numericUpDown16.Value,
                            DefectWeight = (float)numericUpDown15.Value,
                            DefectDensity = checkBox3.Checked ? (float)numericUpDown17.Value : (float)numericUpDown7.Value
                        };

                        dbContext.ProdDefectReport.Add(prodDefectReport);
                        dbContext.SaveChanges();
                    }
                    SaveDefect();
                    dbContext.Dispose();
                }
                // Очистка элементов управления
                comboBox1.Text = default;
                comboBox2.Text = default;
                numericUpDown2.Value = 100;
                numericUpDown4.Value = 1000;
                numericUpDown3.Value = 600;
                numericUpDown5.Value = (decimal) 0.25;
                numericUpDown7.Value = numericUpDown7.Value =  95;
                foreach (ComboBox comboBox in new ComboBox[] { comboBox3, comboBox4, comboBox5, comboBox6, comboBox7, comboBox8, comboBox9 })
                {
                    comboBox.Items.Clear();
                }
                numericUpDown8.Value = numericUpDown9.Value = numericUpDown10.Value = numericUpDown11.Value = numericUpDown12.Value = numericUpDown13.Value =
                    numericUpDown14.Value = numericUpDown15.Value = numericUpDown16.Value = 0;
                checkBox1.Checked = false;
                checkBox2.Checked = false;
                OnLoad();
                
            }

            catch (Exception ex)
            {
                MessageBox.Show($"Произошла ошибка: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                ProjectLogger.LogException("Ошибка в CreateShiftReport", ex);
            }
        }

        public void SaveDefect()
        {
            List<ComboBoxNumericUpDownPair> pairs = new List<ComboBoxNumericUpDownPair>
            {
                new ComboBoxNumericUpDownPair { ComboBox = comboBox3, NumericUpDown = numericUpDown8 },
                new ComboBoxNumericUpDownPair { ComboBox = comboBox4, NumericUpDown = numericUpDown9 },
                new ComboBoxNumericUpDownPair { ComboBox = comboBox5, NumericUpDown = numericUpDown10 },
                new ComboBoxNumericUpDownPair { ComboBox = comboBox6, NumericUpDown = numericUpDown11 },
                new ComboBoxNumericUpDownPair { ComboBox = comboBox7, NumericUpDown = numericUpDown12 },
                new ComboBoxNumericUpDownPair { ComboBox = comboBox8, NumericUpDown = numericUpDown13 },
                new ComboBoxNumericUpDownPair { ComboBox = comboBox9, NumericUpDown = numericUpDown14 }
            };

            using (var dbContext = new ShiftReportDbContext())
            {
                var lastPQReport = dbContext.ProdQualityReport.OrderByDescending(pq => pq.PQReportID).FirstOrDefault();
                foreach (var pair in pairs)
                {
                    if (!string.IsNullOrWhiteSpace(pair.ComboBox.Text))
                    {
                        if (lastPQReport != null)
                        {
                            var defectTypeID = dbContext.DefectTypes
                                .Where(dt => dt.DefectName == pair.ComboBox.Text)
                                .Select(dt => dt.DefectTypeID)
                                .FirstOrDefault();

                            if (defectTypeID != default(int))
                            {
                                var prodDefectReport = new ProdDefectReport
                                {
                                    ProductReport = lastPQReport.PQReportID,
                                    DefectType = defectTypeID,
                                    DefectVolumePack = (float)numericUpDown5.Value,
                                    DefectDensity = checkBox3.Checked ? (float)numericUpDown17.Value : lastPQReport.AvgDensity,
                                    DefectPackCount = (int)pair.NumericUpDown.Value,
                                    DefectVolume = lastPQReport.VolumePack * (int)pair.NumericUpDown.Value,
                                    DefectWeight = (lastPQReport.VolumePack * (int)pair.NumericUpDown.Value) *
                                                  (checkBox3.Checked ? (float)numericUpDown17.Value : lastPQReport.AvgDensity)
                                };
                                dbContext.ProdDefectReport.Add(prodDefectReport);
                                
                            }
                            else
                            {
                                MessageBox.Show($"Не удалось найти соответствующий тип дефекта для {pair.ComboBox.Text} в базе данных.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                        }
                        else
                        {
                            MessageBox.Show("Не удалось найти последнюю запись ProdQualityReport.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }
                dbContext.SaveChanges();
                dbContext.Dispose();
            }
        }


        private void groupBox2_Enter(object sender, EventArgs e)
        {

        }

        private void button6_Click(object sender, EventArgs e)
        {

        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox5_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox6_TextChanged(object sender, EventArgs e)
        {

        }

        private void numericUpDown3_ValueChanged(object sender, EventArgs e)
        {

        }

        private void textBox10_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox11_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox12_TextChanged(object sender, EventArgs e)
        {

        }

        private void numericUpDown8_ValueChanged(object sender, EventArgs e)
        {

        }

        private void textBox19_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox20_TextChanged(object sender, EventArgs e)
        {

        }

        private void comboBox10_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void textBox23_TextChanged(object sender, EventArgs e)
        {

        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {

        }

        private void domainUpDown1_SelectedItemChanged(object sender, EventArgs e)
        {

        }

        private void numericUpDown1_ValueChanged(object sender, EventArgs e)
        {

        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox8_TextChanged(object sender, EventArgs e)
        {

        }

        private void numericUpDown2_ValueChanged(object sender, EventArgs e)
        {

        }

        private void textBox7_TextChanged(object sender, EventArgs e)
        {

        }

        private void numericUpDown4_ValueChanged(object sender, EventArgs e)
        {

        }

        private void textBox9_TextChanged(object sender, EventArgs e)
        {

        }

        private void numericUpDown5_ValueChanged(object sender, EventArgs e)
        {

        }

        private void textBox20_TextChanged_1(object sender, EventArgs e)
        {

        }

        private void numericUpDown6_ValueChanged(object sender, EventArgs e)
        {

        }

        private void numericUpDown7_ValueChanged(object sender, EventArgs e)
        {

        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void textBox13_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox15_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox14_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox18_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox16_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox17_TextChanged(object sender, EventArgs e)
        {

        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            numericUpDown8.Enabled = comboBox3.SelectedIndex >= 0;

            string selectedItem = comboBox3.SelectedItem.ToString();

            // Пройдитесь по остальным комбобоксам и удалите выбранный элемент
            foreach (ComboBox comboBox in new ComboBox[] { comboBox4, comboBox5, comboBox6, comboBox7, comboBox8, comboBox9 })
            {
                comboBox.Items.Remove(selectedItem);
            }
        }

        private void comboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {
            numericUpDown9.Enabled = comboBox4.SelectedIndex >= 0;

            string selectedItem = comboBox4.SelectedItem.ToString();

            // Пройдитесь по остальным комбобоксам и удалите выбранный элемент
            foreach (ComboBox comboBox in new ComboBox[] { comboBox3, comboBox5, comboBox6, comboBox7, comboBox8, comboBox9 })
            {
                comboBox.Items.Remove(selectedItem);
            }
        }

        private void comboBox6_SelectedIndexChanged(object sender, EventArgs e)
        {
            numericUpDown11.Enabled = comboBox6.SelectedIndex >= 0;

            string selectedItem = comboBox6.SelectedItem.ToString();

            // Пройдитесь по остальным комбобоксам и удалите выбранный элемент
            foreach (System.Windows.Forms.ComboBox comboBox in new ComboBox[] { comboBox4, comboBox5, comboBox3, comboBox7, comboBox8, comboBox9 })
            {
                comboBox.Items.Remove(selectedItem);
            }
        }

        private void comboBox5_SelectedIndexChanged(object sender, EventArgs e)
        {
            numericUpDown10.Enabled = comboBox5.SelectedIndex >= 0;

            string selectedItem = comboBox5.SelectedItem.ToString();

            // Пройдитесь по остальным комбобоксам и удалите выбранный элемент
            foreach (ComboBox comboBox in new ComboBox[] { comboBox4, comboBox3, comboBox6, comboBox7, comboBox8, comboBox9 })
            {
                comboBox.Items.Remove(selectedItem);
            }
        }

        private void comboBox9_SelectedIndexChanged(object sender, EventArgs e)
        {
            numericUpDown14.Enabled = comboBox9.SelectedIndex >= 0;

            string selectedItem = comboBox9.SelectedItem.ToString();

            // Пройдитесь по остальным комбобоксам и удалите выбранный элемент
            foreach (ComboBox comboBox in new ComboBox[] { comboBox4, comboBox5, comboBox6, comboBox7, comboBox8, comboBox3 })
            {
                comboBox.Items.Remove(selectedItem);
            }
        }

        private void comboBox7_SelectedIndexChanged(object sender, EventArgs e)
        {
            numericUpDown12.Enabled = comboBox7.SelectedIndex >= 0;

            string selectedItem = comboBox7.SelectedItem.ToString();

            // Пройдитесь по остальным комбобоксам и удалите выбранный элемент
            foreach (ComboBox comboBox in new ComboBox[] { comboBox4, comboBox5, comboBox6, comboBox3, comboBox8, comboBox9 })
            {
                comboBox.Items.Remove(selectedItem);
            }
        }

        private void comboBox8_SelectedIndexChanged(object sender, EventArgs e)
        {
            numericUpDown13.Enabled = comboBox8.SelectedIndex >= 0;

            string selectedItem = comboBox8.SelectedItem.ToString();

            // Пройдитесь по остальным комбобоксам и удалите выбранный элемент
            foreach (ComboBox comboBox in new ComboBox[] { comboBox4, comboBox5, comboBox6, comboBox7, comboBox3, comboBox9 })
            {
                comboBox.Items.Remove(selectedItem);
            }
        }

        private void numericUpDown9_ValueChanged(object sender, EventArgs e)
        {

        }

        private void numericUpDown11_ValueChanged(object sender, EventArgs e)
        {

        }

        private void numericUpDown10_ValueChanged(object sender, EventArgs e)
        {

        }

        private void numericUpDown14_ValueChanged(object sender, EventArgs e)
        {

        }

        private void numericUpDown12_ValueChanged(object sender, EventArgs e)
        {

        }

        private void numericUpDown13_ValueChanged(object sender, EventArgs e)
        {

        }

        private void numericUpDown16_ValueChanged(object sender, EventArgs e)
        {

        }

        private void numericUpDown15_ValueChanged(object sender, EventArgs e)
        {

        }

        private void textBox22_TextChanged(object sender, EventArgs e)
        {

        }

        private void button6_Click_1(object sender, EventArgs e)
        {
            CreateProductShiftReport();
            // Вызываем метод для заполнения выпадающих списков
            FillDropDownLists();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            DateTime now = DateTime.Now;
            DateTime getDate1  = default;
            DateTime getDate2 = default;
            List<int> shiftsDay = new List<int> {};
            List<int> shifts = new List<int> { };
            List<string> stopCategoryes = new List<string> { };
            List<string> prodList = new List<string> { };

            if (now.Hour >= 0 && now.Hour <= 9) { getDate1 = now.AddDays(-1).Date; getDate2 = now.AddDays(-1).Date; shiftsDay.Add(2); }
            else if (now.Hour < 21 && now.Hour > 9) { getDate1 = now.Date; getDate2 = now.Date; shiftsDay.Add(1); }
            if (now.Hour >= 21 && now.Hour < 24) { getDate1 = now.Date; getDate2 = now.Date; shiftsDay.Add(2); }
            int getMethod = 2;
            SaveInXML saver = new SaveInXML();  // Создаем экземпляр класса SaAsDi
            saver.SaveExcelFile(getMethod, prodList, shiftsDay, shifts, stopCategoryes, getDate1, getDate2); // Вызываем метод сохранения
        }

        private void button7_Click(object sender, EventArgs e)
        {
            Form3 form3 = Application.OpenForms.OfType<Form3>().FirstOrDefault();
            form3.Close();
        }

        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {
            numericUpDown17.Enabled = checkBox3.Checked;
        }

        private void groupBox11_Enter(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            SmallGrid();
        }

        private void numericUpDown17_ValueChanged(object sender, EventArgs e)
        {

        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void textBox22_TextChanged_1(object sender, EventArgs e)
        {

        }

        private void dataGridView1_CellContentClick_1(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            string getText = comboBox1.Text;
            Form8 form8 = new Form8(getText);
            form8.Show();
        }

        private void Form3_Load(object sender, EventArgs e)
        {

        }
    }
}
