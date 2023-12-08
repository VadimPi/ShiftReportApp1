using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ShiftReportApp1
{
    public partial class Form6 : Form
    {
        private PassManager passManager;

        public Form6()
        {
            InitializeComponent();
        }
        private void FillTable(int getMethod)
        {
            DateTime varDate1 = (DateTime)dateTimePicker1.Value;
            DateTime varDate2 = (DateTime)dateTimePicker1.Value;
            List<int> shiftDays = new List<int> ();
            if (domainUpDown1.Text == "День") shiftDays = new List<int> { 1 };
            else if (domainUpDown1.Text == "Ночь") shiftDays = new List<int> { 2 };
            List<int> shifts = new List<int> { 1, 2, 3, 4};
            List<string> stopCategoryes = new List<string> ();
            LINQRequest newReport = new LINQRequest();
            DataTable dataTable = newReport.ExtractProduct(getMethod, varDate1, varDate2, shiftDays, shifts, stopCategoryes);
            dataGridView1.DataSource = dataTable;

        }

        private void dataGridView1_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {
            if (radioButton1.Checked || radioButton2.Checked || radioButton6.Checked)
            {
                // Проверить, что редактируется существующая строка, а не строка для нового элемента
                if (e.RowIndex < dataGridView1.Rows.Count - 1)
                {
                    // Разрешить редактирование
                    e.Cancel = false;
                }
                else
                {
                    // Запретить редактирование новой строки
                    e.Cancel = true;
                }
            }
        }

        public void SaveChangesToDatabase(DataGridView dataGridView)
        {
            try
            {
                ProjectLogger.LogDebug("Начало ExtractProduct (SaveChangesToDatabase)");
                if (radioButton1.Checked)
                {
                    using (var dbContext = new ShiftReportDbContext())
                    {
                        foreach (DataGridViewRow row in dataGridView.Rows)
                        {
                            // Пропускаем строки, которые не содержат данных
                            if (row.IsNewRow || row.Cells[0].Value == null)
                                continue;

                            int numProdReport = Convert.ToInt32(row.Cells["# записи продукта"].Value);

                            // Ищем запись в базе данных по номеру отчета
                            var existingRecord = dbContext.ProdQualityReport.FirstOrDefault(pr => pr.PQReportID == numProdReport);

                            if (existingRecord != null)
                            {
                                string productName = row.Cells["Марка"].Value.ToString();
                                var product = dbContext.ProductCategories.FirstOrDefault(pc => pc.ProductName == productName);
                                // Обновляем поля записи
                                existingRecord.Product = product.ProductID;
                                existingRecord.Unspecified = Convert.ToBoolean(row.Cells["Неуказанная плт-ть"].Value);
                                existingRecord.Regarding = Convert.ToBoolean(row.Cells["Пересорт"].Value);
                                existingRecord.ProdLength = Convert.ToInt32(row.Cells["Длинна"].Value);
                                existingRecord.ProdWidth = Convert.ToInt32(row.Cells["Ширина"].Value);
                                existingRecord.ProdDepth = Convert.ToInt32(row.Cells["Толщина"].Value);
                                existingRecord.AvgDensity = Convert.ToSingle(row.Cells["Ср. плотность"].Value);
                                existingRecord.VolumePack = Convert.ToSingle(row.Cells["Объем пачки"].Value);
                                existingRecord.PackCount = Convert.ToInt32(row.Cells["Кол-во пачек"].Value);
                                existingRecord.Weight = Convert.ToSingle(row.Cells["Вес"].Value);
                                existingRecord.VolumeProduct = Convert.ToSingle(row.Cells["Объем"].Value);
                                // Обновление для других полей аналогично
                            }
                            // Аналогичные блоки кода добавьте для других отчетов
                        }
                        // Сохраняем изменения в базе данных
                        dbContext.SaveChanges();
                        MessageBox.Show("Данные успешно сохранены", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        dbContext.Dispose();
                    }
                }
                //---------------------------------------------------------------------------------------------------------------------------------------
                if (radioButton2.Checked)
                {
                    using (var dbContext = new ShiftReportDbContext())
                    {
                        foreach (DataGridViewRow row in dataGridView.Rows)
                        {
                            // Пропускаем строки, которые не содержат данных
                            if (row.IsNewRow || row.Cells[0].Value == null)
                                continue;

                            int numReport = Convert.ToInt32(row.Cells["# записи остановки"].Value);

                            // Ищем запись в базе данных по номеру отчета
                            var existingRecord = dbContext.StopsReport.FirstOrDefault(str => str.StopReportID == numReport);

                            if (existingRecord != null)
                            {
                                string stopName = row.Cells["Название остановки"].Value.ToString();
                                var stop = dbContext.StopType.FirstOrDefault(st => st.StopName == stopName);
                                string placeName = row.Cells["Место остановки"].Value.ToString();
                                var place = dbContext.PlaceInLine.FirstOrDefault(pil => pil.PlacesName == placeName);
                                // Обновляем поля записи
                                existingRecord.StopType = stop.StopTypeID;
                                existingRecord.PlaceStop = place.PlacesID;
                                existingRecord.StopFirstTime = Convert.ToString(row.Cells["Начало остановки"].Value);
                                existingRecord.StopEndTime = Convert.ToString(row.Cells["Конец остановки"].Value);
                                existingRecord.DurationStopMin = Convert.ToInt32(row.Cells["Длительность"].Value);
                                existingRecord.BreakdownWithoutStop = Convert.ToBoolean(row.Cells["Остановка выпуска"].Value);
                                existingRecord.CommentStop = Convert.ToString(row.Cells["Комментарий"].Value);
                                existingRecord.Centrifuge = Convert.ToInt32(row.Cells["Фуга"].Value);
                                // Обновление для других полей аналогично
                            }
                            // Аналогичные блоки кода добавьте для других отчетов
                        }
                        // Сохраняем изменения в базе данных
                        dbContext.SaveChanges();
                        MessageBox.Show("Данные успешно сохранены", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        dbContext.Dispose();
                    }
                }
                //---------------------------------------------------------------------------------------------------------------------------------------
                if (radioButton6.Checked)
                {
                    using (var dbContext = new ShiftReportDbContext())
                    {
                        foreach (DataGridViewRow row in dataGridView.Rows)
                        {
                            // Пропускаем строки, которые не содержат данных
                            if (row.IsNewRow || row.Cells[0].Value == null)
                                continue;

                            int numReport = Convert.ToInt32(row.Cells["# записи дефекта"].Value);

                            // Ищем запись в базе данных по номеру отчета
                            var existingRecord = dbContext.ProdDefectReport.FirstOrDefault(pdr => pdr.DefectReportID == numReport);

                            if (existingRecord != null)
                            {
                                string defectName = row.Cells["Дефект"].Value.ToString();
                                var defect = dbContext.DefectTypes.FirstOrDefault(dt => dt.DefectName == defectName);

                                existingRecord.ProductReport = Convert.ToInt32(row.Cells["# записи продукта"].Value);
                                existingRecord.DefectType = defect.DefectTypeID;
                                existingRecord.DefectVolumePack = Convert.ToSingle(row.Cells["Объем пачки"].Value);
                                existingRecord.DefectDensity = Convert.ToSingle(row.Cells["Плотность"].Value);
                                existingRecord.DefectPackCount = Convert.ToInt32(row.Cells["Кол-во пачек"].Value);
                                existingRecord.DefectVolume = Convert.ToSingle(row.Cells["Объем"].Value);
                                existingRecord.DefectWeight = Convert.ToSingle(row.Cells["Вес"].Value);
                            }
                            // Аналогичные блоки кода добавьте для других отчетов
                        }
                        // Сохраняем изменения в базе данных
                        dbContext.SaveChanges();
                        MessageBox.Show("Данные успешно сохранены", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        dbContext.Dispose();
                    }
                }
                //---------------------------------------------------------------------------------------------------------------------------------------
                if (radioButton3.Checked)
                {
                    using (var dbContext = new ShiftReportDbContext())
                    {
                        foreach (DataGridViewRow row in dataGridView.Rows)
                        {
                            // Пропускаем строки, которые не содержат данных
                            if (row.IsNewRow || row.Cells[0].Value == null)
                                continue;

                            int ProductID = -1;
                            if (row.Cells["ID"].Value != DBNull.Value) ProductID = Convert.ToInt32(row.Cells["ID"].Value);
                            string ProductName = Convert.ToString(row.Cells["Имя продукта"].Value);

                            // Ищем запись в базе данных по номеру отчета
                            var existingRecord = dbContext.ProductCategories.FirstOrDefault(pc => pc.ProductID == ProductID);

                            if (existingRecord != null)
                            {
                                existingRecord.ProductName = Convert.ToString(row.Cells["Имя продукта"].Value);
                                existingRecord.DensityMin = Convert.ToSingle(row.Cells["Минимальная плотность"].Value);
                                existingRecord.DensityMax = Convert.ToSingle(row.Cells["Максимальная плотность"].Value);
                            }
                            // Аналогичные блоки кода добавьте для других отчетов
                            if (ProductID == -1)
                            {
                                var newRow = new ProductCat
                                {
                                    ProductName = Convert.ToString(row.Cells["Имя продукта"].Value),
                                    DensityMin = Convert.ToSingle(row.Cells["Минимальная плотность"].Value),
                                    DensityMax = Convert.ToSingle(row.Cells["Максимальная плотность"].Value)
                                };

                                dbContext.ProductCategories.Add(newRow);
                            }
                        }
                        // Сохраняем изменения в базе данных
                        dbContext.SaveChanges();
                        MessageBox.Show("Данные успешно сохранены", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        dbContext.Dispose();
                    }
                }
                //---------------------------------------------------------------------------------------------------------------------------------------
                if (radioButton4.Checked)
                {
                    using (var dbContext = new ShiftReportDbContext())
                    {
                        foreach (DataGridViewRow row in dataGridView.Rows)
                        {
                            // Пропускаем строки, которые не содержат данных
                            if (row.IsNewRow || row.Cells[0].Value == null)
                                continue;

                            int PlacesID = -1;
                            if (row.Cells["ID"].Value != DBNull.Value) PlacesID = Convert.ToInt32(row.Cells["ID"].Value);
                            string PlacesName = Convert.ToString(row.Cells["Место"].Value);

                            // Ищем запись в базе данных по номеру отчета
                            var existingRecord = dbContext.PlaceInLine.FirstOrDefault(pil => pil.PlacesID == PlacesID);

                            if (existingRecord != null)
                            {
                                existingRecord.PlacesName = Convert.ToString(row.Cells["Место"].Value);
                                existingRecord.Section = Convert.ToString(row.Cells["Машина"].Value);
                            }
                            // Аналогичные блоки кода добавьте для других отчетов
                            if (PlacesID == -1)
                            {
                                var newRow = new PlaceInLine
                                {
                                    PlacesName = Convert.ToString(row.Cells["Место"].Value),
                                    Section = Convert.ToString(row.Cells["Машина"].Value)
                                };

                                dbContext.PlaceInLine.Add(newRow);
                            }
                        }
                        // Сохраняем изменения в базе данных
                        dbContext.SaveChanges();
                        MessageBox.Show("Данные успешно сохранены", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        dbContext.Dispose();
                    }
                }
                //---------------------------------------------------------------------------------------------------------------------------------------
                if (radioButton5.Checked)
                {
                    using (var dbContext = new ShiftReportDbContext())
                    {
                        foreach (DataGridViewRow row in dataGridView.Rows)
                        {
                            // Пропускаем строки, которые не содержат данных
                            if (row.IsNewRow || row.Cells[0].Value == null)
                                continue;
                            int StopID = -1;
                            if (row.Cells["ID"].Value != DBNull.Value) StopID = Convert.ToInt32(row.Cells["ID"].Value);
                            string StopName = Convert.ToString(row.Cells["Остановка"].Value);

                            // Ищем запись в базе данных по номеру отчета
                            var existingRecord = dbContext.StopType.FirstOrDefault(st => st.StopTypeID == StopID);

                            if (existingRecord != null)
                            {
                                existingRecord.StopName = Convert.ToString(row.Cells["Остановка"].Value);
                                existingRecord.StopCategory = Convert.ToString(row.Cells["Категория"].Value);
                            }
                            // Аналогичные блоки кода добавьте для других отчетов
                            if (StopID == -1)
                            {
                                var newRow = new StopType
                                {
                                    StopName = Convert.ToString(row.Cells["Остановка"].Value),
                                    StopCategory = Convert.ToString(row.Cells["Категория"].Value)
                                };

                                dbContext.StopType.Add(newRow);
                            }
                        }
                        // Сохраняем изменения в базе данных
                        dbContext.SaveChanges();
                        MessageBox.Show("Данные успешно сохранены", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        dbContext.Dispose();
                    }
                }
                //---------------------------------------------------------------------------------------------------------------------------------------
                if (radioButton7.Checked)
                {
                    using (var dbContext = new ShiftReportDbContext())
                    {
                        foreach (DataGridViewRow row in dataGridView.Rows)
                        {
                            // Пропускаем строки, которые не содержат данных
                            if (row.IsNewRow || row.Cells[0].Value == null)
                                continue;

                            int DefectID = -1;
                            if (row.Cells["ID"].Value != DBNull.Value) DefectID = Convert.ToInt32(row.Cells["ID"].Value);
                            string DefectName = Convert.ToString(row.Cells["Имя дефекта"].Value);

                            // Ищем запись в базе данных по номеру отчета
                            var existingRecord = dbContext.DefectTypes.FirstOrDefault(dt => dt.DefectTypeID == DefectID);

                            if (existingRecord != null)
                            {
                                existingRecord.DefectName = Convert.ToString(row.Cells["Имя дефекта"].Value);
                            }
                            // Аналогичные блоки кода добавьте для других отчетов
                            if (DefectID == -1)
                            {
                                var newRow = new DefectType
                                {
                                    DefectName = Convert.ToString(row.Cells["Имя дефекта"].Value)
                                };

                                dbContext.DefectTypes.Add(newRow);
                            }
                        }
                        // Сохраняем изменения в базе данных
                        dbContext.SaveChanges();
                        MessageBox.Show("Данные успешно сохранены", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        dbContext.Dispose();
                    }
                }
                // исключения
                ProjectLogger.LogDebug("Конец ExtractProduct (SaveChangesToDatabase)");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                ProjectLogger.LogException("Ошибка в ExtractProduct (SaveChangesToDatabase)", ex);
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            Form6 form6 = Application.OpenForms.OfType<Form6>().FirstOrDefault();
            form6.Close();
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            passManager.SetPassword(textBox2.Text);
        }

        private void radioButton5_CheckedChanged(object sender, EventArgs e)
        {
            dataGridView1.AllowUserToAddRows = true;
            FillTable(35);
        }

        private void radioButton4_CheckedChanged(object sender, EventArgs e)
        {
            dataGridView1.AllowUserToAddRows = true;
            FillTable(36);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Form9 form9 = new Form9();
            form9.Show();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            SaveChangesToDatabase(dataGridView1);
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            dataGridView1.AllowUserToAddRows = false;
            FillTable(32);
        }

        private void radioButton6_CheckedChanged(object sender, EventArgs e)
        {
            dataGridView1.AllowUserToAddRows = false;
            FillTable(31);
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            dataGridView1.AllowUserToAddRows = false;
            FillTable(30);
        }

        private void radioButton7_CheckedChanged(object sender, EventArgs e)
        {
            dataGridView1.AllowUserToAddRows = true;
            FillTable(34);
        }

        private void radioButton3_CheckedChanged(object sender, EventArgs e)
        {
            dataGridView1.AllowUserToAddRows = true;
            FillTable(33);
        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {

        }

        private void domainUpDown1_SelectedItemChanged(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox8_TextChanged(object sender, EventArgs e)
        {

        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void Form6_Load(object sender, EventArgs e)
        {

        }
    }
}
