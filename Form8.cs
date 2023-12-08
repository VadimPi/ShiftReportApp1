using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.EMMA;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Font = System.Drawing.Font;

namespace ShiftReportApp1
{
    public partial class Form8 : Form
    {
        public Form8(string getString)
        {
            InitializeComponent();
            fillGird(getString);
            dataGridView1.ColumnHeadersDefaultCellStyle.Font = new Font("Arial", 12); // Устанавливаем шрифт и размер
            dataGridView1.DefaultCellStyle.Font = new Font("Arial", 12); // Устанавливаем шрифт и размер
            dataGridView1.RowTemplate.Height = 20; // Устанавливаем высоту строки
            dataGridView1.ColumnHeadersHeight = 20; // Устанавливаем высоту заголовка столбца
            dataGridView1.CellBorderStyle = DataGridViewCellBorderStyle.Single;
            if (dataGridView1.Rows.Count > 0)
            {
                dataGridView1.Columns[0].Width = 120;
                dataGridView1.Columns[1].Width = 70;
                dataGridView1.Columns[2].Width = 70;
                dataGridView1.Columns[3].Width = 70;
                dataGridView1.Columns[4].Width = 80;
                dataGridView1.Columns[5].Width = 80;
                dataGridView1.Columns[6].Width = 80;
                dataGridView1.Columns[7].Width = 80;
                dataGridView1.Columns[8].Width = 80;
                dataGridView1.Columns[9].Width = 80;
                dataGridView1.Columns[10].Width = 100;
                dataGridView1.Columns[11].Width = 80;
                dataGridView1.Columns[12].Width = 80;
                dataGridView1.Columns[13].Width = 80;
                dataGridView1.Columns[14].Width = 80;
                dataGridView1.Columns[15].Width = 80;
                dataGridView1.Columns[16].Width = 80;
                //dataGridView1.Columns[17].Width = 80;
            }
        }

        private void Form8_Load(object sender, EventArgs e)
        {

        }

        private void fillGird(string getString)
        {
            var reportNumber = 0;

            if (!string.IsNullOrWhiteSpace(getString) && int.TryParse(getString.Split(' ')[0], out reportNumber))
            {
                using (var dbContext = new ShiftReportDbContext())
                {
                    var prodQualityReports =
                    from pqr in dbContext.ProdQualityReport
                    join pc in dbContext.ProductCategories on pqr.Product equals pc.ProductID
                    join sr in dbContext.ShiftReport on pqr.Report equals sr.ShiftReportID
                    join pdr in dbContext.ProdDefectReport on pqr.PQReportID equals pdr.ProductReport into defectReports
                    from pdr in defectReports.DefaultIfEmpty() // Выполняем LEFT JOIN
                    where sr.ShiftReportID == reportNumber
                    group new { pqr, pc, pdr } by new
                    {
                        pc.ProductName,
                        pqr.Unspecified,
                        pqr.ProdDepth,
                        pqr.ProdLength,
                        pqr.ProdWidth,
                        pqr.Regarding
                    } into grp
                    select new
                    {
                        ProductName = grp.Key.ProductName,
                        Unspecified = grp.Key.Unspecified,
                        ProdDepth = grp.Key.ProdDepth,
                        Length = grp.Key.ProdLength,
                        Width = grp.Key.ProdWidth,
                        Regarding = grp.Key.Regarding,
                        AvgDensity = grp.Average(x => x.pqr.AvgDensity),
                        PackCount = grp.Average(x => x.pqr.PackCount),
                        Volume = Math.Round(grp.Average(x => x.pqr.VolumeProduct), 3),
                        Weight = Math.Round(grp.Average(x => x.pqr.Weight)),
                        LowQualCount = grp.Sum(x => x.pdr.DefectType >= 1 && x.pdr.DefectType <= 7 ? x.pdr.DefectPackCount : 0),
                        LowQualVol = Math.Round(grp.Sum(x => x.pdr.DefectType >= 1 && x.pdr.DefectType <= 7 ? x.pdr.DefectVolume : 0),3),
                        LowQualWeight = Math.Round(grp.Sum(x => x.pdr.DefectType >= 1 && x.pdr.DefectType <= 7 ? x.pdr.DefectWeight : 0)),
                        RejectVol = Math.Round(grp.Sum(x => x.pdr.DefectType == 8 ? x.pdr.DefectVolume : 0),3),
                        RejectWeight = Math.Round(grp.Sum(x => x.pdr.DefectType == 8 ? x.pdr.DefectWeight : 0)),
                        PercentLowQual = Math.Round((grp.Sum(x => x.pdr.DefectType >= 1 && x.pdr.DefectType <= 7 ? x.pdr.DefectWeight : 0) / grp.Average(x => x.pqr.Weight)) * 100.0, 3),
                        PercentReject = Math.Round((grp.Sum(x => x.pdr.DefectType == 8 ? x.pdr.DefectWeight : 0) / grp.Average(x => x.pqr.Weight)) * 100.0, 2),
                                            };
                    prodQualityReports.ToList();

                    DataTable dataTable = new DataTable();
                    dataTable.Columns.Add("Марка", typeof(string));
                    dataTable.Columns.Add("Толщ.", typeof(int));
                    dataTable.Columns.Add("Длин.", typeof(int));
                    dataTable.Columns.Add("Шир.", typeof(int));
                    dataTable.Columns.Add("Неуказ.", typeof(bool));
                    dataTable.Columns.Add("Пересорт", typeof(bool));
                    dataTable.Columns.Add("Плот-ть", typeof(float));
                    dataTable.Columns.Add("Кол-во пачек", typeof(int));
                    dataTable.Columns.Add("Объем", typeof(float));
                    dataTable.Columns.Add("Вес", typeof(float));
                    dataTable.Columns.Add("Кол-во пачек ОС", typeof(int));
                    dataTable.Columns.Add("Объем ОС", typeof(float));
                    dataTable.Columns.Add("Вес ОС", typeof(float));
                    dataTable.Columns.Add("Объем обрезь", typeof(float));
                    dataTable.Columns.Add("Вес обрезь", typeof(float));
                    dataTable.Columns.Add("Процент ОС", typeof(double));
                    dataTable.Columns.Add("Процент обрезь", typeof(double));
                    //dataTable.Columns.Add("Процент пересорта", typeof(double));

                    // Заполняем DataTable
                    foreach (var report in prodQualityReports)
                    {
                        DataRow row = dataTable.NewRow();
                        row["Марка"] = report.ProductName;
                        row["Толщ."] = report.ProdDepth;
                        row["Длин."] = report.Length;
                        row["Шир."] = report.Width;
                        row["Неуказ."] = report.Unspecified;
                        row["Пересорт"] = report.Regarding;
                        row["Плот-ть"] = report.AvgDensity;
                        row["Кол-во пачек"] = report.PackCount;
                        row["Объем"] = report.Volume;
                        row["Вес"] = report.Weight;
                        row["Кол-во пачек ОС"] = report.LowQualCount;
                        row["Объем ОС"] = report.LowQualVol;
                        row["Вес ОС"] = report.LowQualWeight;
                        row["Объем обрезь"] = report.RejectVol;
                        row["Вес обрезь"] = report.RejectWeight;
                        row["Процент ОС"] = report.PercentLowQual;
                        row["Процент обрезь"] = report.PercentReject;
                        //row["Процент пересорта"] = report.PercentRegarding;
                        // Заполняйте остальные поля аналогично
                        dataTable.Rows.Add(row);

                        dataGridView1.DataSource = dataTable;
                        dataGridView1.CellFormatting += dataGridView1_CellFormatting;
                    }
                    dbContext.Dispose();
                }
            }
            else
            {
                dataGridView1.Text = "Неверный формат отчета.";
            }
        }
        private void dataGridView1_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            for (int columnIndex = 7; columnIndex <= 14; columnIndex++)
            {
                decimal sum = 0;

                for (int rowIndex = 0; rowIndex < dataGridView1.Rows.Count - 1; rowIndex++)
                {
                    if (dataGridView1[columnIndex, rowIndex].Value != null &&
                        decimal.TryParse(dataGridView1[columnIndex, rowIndex].Value.ToString(), out decimal cellValue))
                    {
                        sum += cellValue;
                    }
                }

                // Отображение суммы в нужной ячейке
                dataGridView1[columnIndex, dataGridView1.Rows.Count - 1].Value = sum;
            }
            int weightColumnIndex = 9;
            int osWeightColumnIndex = 12;
            int rejWeightColumnIndex = 14;
            int osPercentColumnIndex = 15;
            int rejPercentColumnIndex = 16;
            int endRowIndex = dataGridView1.Rows.Count - 1;
            decimal.TryParse(dataGridView1[osWeightColumnIndex, endRowIndex].Value.ToString(), out decimal osValue);
            decimal.TryParse(dataGridView1[rejWeightColumnIndex, endRowIndex].Value.ToString(), out decimal rejValue);
            decimal.TryParse(dataGridView1[weightColumnIndex, endRowIndex].Value.ToString(), out decimal weightValue);
            dataGridView1[osPercentColumnIndex, endRowIndex].Value = Math.Round(osValue / weightValue * 100,2);
            dataGridView1[rejPercentColumnIndex, endRowIndex].Value = Math.Round(rejValue / weightValue * 100,2);
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
    }
}
