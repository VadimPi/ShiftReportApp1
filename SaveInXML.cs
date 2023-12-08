using ClosedXML.Excel;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Spreadsheet;
using Npgsql;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.StartPanel;
using Path = System.IO.Path;

namespace ShiftReportApp1
{
    internal class SaveInXML
    {
        public string FilePath { get; set; }
        public void SaveExcelFile(int getMethod, List<string> prodList, List<int> shiftsDays, List<int> shifts, List<string> stopCategoryes,
             DateTime varDate1 = default, DateTime varDate2 = default)
        {
            try
            {
                ProjectLogger.LogDebug("Начало SaveExcelFile (SaveIn)");
                LINQRequest newReport = new LINQRequest();

                DataTable dataTable = new DataTable();
                List <DataTable> dataTablePivot = new List<DataTable> { new DataTable(), new DataTable() };
                if (getMethod != 21) { dataTable = newReport.ExtractProduct(getMethod, varDate1, varDate2, shiftsDays, shifts, stopCategoryes); }
                else {dataTablePivot = newReport.ExtractProductList(getMethod, varDate1, varDate2, shiftsDays, shifts, stopCategoryes); }
                var workbook = new XLWorkbook();

                using (SaveFileDialog saveFileDialog = new SaveFileDialog())
                {
                    saveFileDialog.Filter = "Файлы Excel (*.xlsx)|*.xlsx|Все файлы (*.*)|*.*";
                    saveFileDialog.FilterIndex = 1;
                    saveFileDialog.RestoreDirectory = true;

                    DialogResult result = saveFileDialog.ShowDialog();

                    if (result == DialogResult.OK)
                    {
                        string filePath = saveFileDialog.FileName;
                        string templateFilePath = "";
                        // Копируем файл шаблона в папку с исполняемым файлом
                        if (getMethod == 0 || getMethod == 3)
                        {
                            templateFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "qtemp2.xlsx");
                        }
                        else if ( getMethod == 2)
                        {
                            templateFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "qtemp1.xlsx");
                        }
                        else if (getMethod == 4 || getMethod == 5)
                        {
                            templateFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "qtemp3.xlsx");
                        }
                        else if (getMethod >= 10 && getMethod <= 15)
                        {
                            templateFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "qtemp4.xlsx");
                        }
                        else if (getMethod == 21)
                        { templateFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "qtemp5.xlsx"); }

                        FilePath = filePath;
                        File.Copy(templateFilePath, Path.Combine(Path.GetDirectoryName(filePath), Path.GetFileNameWithoutExtension(filePath) + ".xlsx"), true);
                        workbook = new XLWorkbook(Path.Combine(Path.GetDirectoryName(filePath), Path.GetFileNameWithoutExtension(filePath) + ".xlsx"));
                    }
                    else if (result == DialogResult.Cancel)
                    {
                        // Пользователь отменил операцию сохранения, завершаем метод
                        ProjectLogger.LogDebug("Операция сохранения отменена пользователем");
                        return;
                    }
                }
                if (getMethod == 2)
                {
                    WriteToExcel24hours(workbook, dataTable, getMethod, prodList, shiftsDays, shifts, stopCategoryes, varDate1, varDate2);
                }
                else if (getMethod == 0 || getMethod == 3)
                {
                    WriteToExcelInToDates(workbook, dataTable, getMethod, prodList, shiftsDays, shifts, stopCategoryes, varDate1, varDate2);
                }
                else if (getMethod == 4 || getMethod == 5)
                {
                    WriteToExcelBrokenProduct(workbook, dataTable, getMethod, shiftsDays, shifts, stopCategoryes, varDate1, varDate2);
                }
                else if (getMethod == 21)
                {
                    WriteToExcelPivote(workbook, dataTablePivot, getMethod, shiftsDays, shifts, stopCategoryes, varDate1, varDate2);
                }
                MessageBox.Show($"Файл сохранен", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                ProjectLogger.LogDebug("Конец SaveExcelFile (SaveIn)");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                ProjectLogger.LogException("Ошибка в SaveExcelFile (SaveIn)", ex);
            }
        }
        //  -----------------------------------------------------------------------------------------
        // Заполнение excel за предыдущие сутки
        public void WriteToExcel24hours(XLWorkbook workbook, DataTable dataTable, int getMethod, List<string> prodList, List<int> shiftsDays, List<int> shifts, List<string> stopCategoryes,
             DateTime varDate1 = default, DateTime varDate2 = default)
        {
            try
            {
                ProjectLogger.LogDebug("Начало метода WriteToExcel24hours");
                var worksheet = workbook.Worksheet(1);
                int firstRow = 0, secondRow = 0, countRow1 = 0, countRow2 = 0, countRow = 0;
                int shiftNumConst = (int)dataTable.Rows[0][1];

                DateTime shiftDate = (DateTime)dataTable.Rows[0][0];
                DateTime shiftDate2 = shiftDate.AddDays(1);
                shiftDate.ToShortDateString();
                shiftDate2.ToShortDateString();
                worksheet.Cell(63, 11).Value = shiftDate2;

                int sumPack1 = 0, sumPack2 = 0, sumOS1 = 0, sumOS2 = 0;
                float sumVolume1 = 0, sumVolume2 = 0, sumWeight1 = 0, sumWeight2 = 0, sumVolumeOS1 = 0, sumVolumeOS2 = 0, sumWeightOS1 = 0, sumWeightOS2 = 0;
                float sumVolumeRej1 = 0, sumVolumeRej2 = 0, sumWeightRej1 = 0, sumWeightRej2 = 0, sumRegarding1 = 0, sumRegarding2 = 0;

                for (int row = 0; row < dataTable.Rows.Count; row++) // блок заполнения смен
                {
                    if (shiftNumConst == (int)dataTable.Rows[row][1]) // начало записи
                    {
                        countRow1++;
                        if ((bool)dataTable.Rows[row][17]) // подсчет переорта см 1
                        {
                            sumRegarding1 = (float)(sumRegarding1 + Math.Round((float)dataTable.Rows[row][11]));
                        }
                    }
                    else
                    {
                        countRow2++;
                        if ((bool)dataTable.Rows[row][17]) // подсчет переорта см 2
                        {
                            sumRegarding2 = (float)(sumRegarding1 + Math.Round((float)dataTable.Rows[row][11]));
                        }
                    }
                    if (shiftsDays.Count == 2) // подсчет за сутки
                    {
                        worksheet.Cell(2, 7).Value = worksheet.Cell(31, 7).Value = worksheet.Cell(63, 8).Value = shiftDate; // дата
                        firstRow = shiftNumConst == (int)dataTable.Rows[row][1] ? 6 : 35; // выбор первой строки сменных от смены
                        secondRow = shiftNumConst == (int)dataTable.Rows[row][1] ? 66 : 81; // выбор первой строки суточного от смены
                        countRow = shiftNumConst == (int)dataTable.Rows[row][1] ? countRow1 : countRow2; // счет продуктов по сменам
                    }
                    else if (shiftsDays.Count == 1 && shiftsDays[0] == 1) // предыдущая смена дневная
                    {

                        worksheet.Cell(2, 7).Value = worksheet.Cell(63, 8).Value = shiftDate; // дата
                        if (shiftNumConst != (int)dataTable.Rows[row][1]) // прерывание если записали две смены
                        {
                            return;
                        }
                        firstRow = 6;
                        secondRow = 66;
                        countRow = shiftNumConst == (int)dataTable.Rows[row][1] ? countRow1 : countRow2;// счет продуктов по сменам
                    }
                    else if (shiftsDays.Count == 1 && shiftsDays[0] == 2) // предыдущая смена ночная
                    {

                        worksheet.Cell(31, 7).Value = worksheet.Cell(63, 8).Value = shiftDate; // дата
                        if (shiftNumConst != (int)dataTable.Rows[row][1]) // прерывание если записали две смены
                        {
                            return;
                        }
                        firstRow = 35;
                        secondRow = 81;
                        countRow = shiftNumConst == (int)dataTable.Rows[row][1] ? countRow1 : countRow2;// счет продуктов по сменам
                    }

                    worksheet.Cell(countRow + firstRow, 1).Value = worksheet.Cell(countRow + secondRow, 1).Value = countRow; // порядковый номер продукта (в Кандыгаше №партии
                    worksheet.Cell(firstRow - 4, 16).Value = worksheet.Cell(countRow + secondRow, 2).Value = (int)dataTable.Rows[row][1]; // смена
                    worksheet.Cell(countRow + firstRow, 2).Value = worksheet.Cell(countRow + secondRow, 3).Value = GetProductNameDay(dataTable, row); // наименование продукта
                    worksheet.Cell(countRow + firstRow, 4).Value = worksheet.Cell(countRow + secondRow, 10).Value = (float)dataTable.Rows[row][8]; // Объем 1 пачки
                    worksheet.Cell(countRow + firstRow, 5).Value = worksheet.Cell(countRow + secondRow, 4).Value = (float)dataTable.Rows[row][7]; // средняя плотность
                    worksheet.Cell(countRow + firstRow, 8).Value = worksheet.Cell(countRow + secondRow, 11).Value = (int)dataTable.Rows[row][9]; // пачек гп
                    worksheet.Cell(countRow + firstRow, 10).Value = worksheet.Cell(countRow + secondRow, 12).Value = Math.Round((float)dataTable.Rows[row][10], 3); // объем
                    worksheet.Cell(countRow + firstRow, 11).Value = worksheet.Cell(countRow + secondRow, 13).Value = Math.Round((float)dataTable.Rows[row][11]); // вес
                    worksheet.Cell(countRow + firstRow, 12).Value = worksheet.Cell(countRow + secondRow, 6).Value = (float)dataTable.Rows[row][18]; // плотность ос
                    worksheet.Cell(countRow + firstRow, 13).Value = worksheet.Cell(countRow + secondRow, 15).Value = (int)dataTable.Rows[row][12]; // пачек ос
                    worksheet.Cell(countRow + firstRow, 15).Value = worksheet.Cell(countRow + secondRow, 16).Value = Math.Round((float)dataTable.Rows[row][13], 3); // Объем ос
                    worksheet.Cell(countRow + firstRow, 17).Value = worksheet.Cell(countRow + secondRow, 18).Value = Math.Round((float)dataTable.Rows[row][14]); // вес ос
                    worksheet.Cell(countRow + firstRow, 18).Value = worksheet.Cell(countRow + secondRow, 8).Value = (float)dataTable.Rows[row][18]; // плотность обрези
                    worksheet.Cell(countRow + firstRow, 19).Value = worksheet.Cell(countRow + secondRow, 19).Value = Math.Round((float)dataTable.Rows[row][15], 3); // Объем обрези
                    worksheet.Cell(countRow + firstRow, 20).Value = worksheet.Cell(countRow + secondRow, 20).Value = Math.Round((float)dataTable.Rows[row][16]); // вес обрези

                    if (shiftNumConst == (int)dataTable.Rows[row][1]) // суммирование
                    {
                        if (shiftsDays.Count == 2 || (shiftsDays.Count == 1 && shiftsDays[0] == 1))
                        {
                            sumPack1 += (int)dataTable.Rows[row][9];
                            sumVolume1 = (float)(sumVolume1 + Math.Round((float)dataTable.Rows[row][10], 3));
                            sumWeight1 = (float)(sumWeight1 + Math.Round((float)dataTable.Rows[row][11]));
                            sumOS1 += (int)dataTable.Rows[row][12];
                            sumVolumeOS1 = (float)(sumVolumeOS1 + Math.Round((float)dataTable.Rows[row][13], 3));
                            sumWeightOS1 = (float)(sumWeightOS1 + Math.Round((float)dataTable.Rows[row][14]));
                            sumVolumeRej1 = (float)(sumVolumeRej1 + Math.Round((float)dataTable.Rows[row][15], 3));
                            sumWeightRej1 = (float)(sumWeightRej1 + Math.Round((float)dataTable.Rows[row][16]));

                            if ((bool)dataTable.Rows[row][17])
                            {
                                sumRegarding1 = (float)(sumRegarding1 + Math.Round((float)dataTable.Rows[row][11]));
                            }
                        }
                        else if (shiftsDays.Count == 1 && shiftsDays[0] == 2)
                        {
                            sumPack2 += (int)dataTable.Rows[row][9];
                            sumVolume2 = (float)(sumVolume2 + Math.Round((float)dataTable.Rows[row][10], 3));
                            sumWeight2 = (float)(sumWeight2 + Math.Round((float)dataTable.Rows[row][11]));
                            sumOS2 += (int)dataTable.Rows[row][12];
                            sumVolumeOS2 = (float)(sumVolumeOS2 + Math.Round((float)dataTable.Rows[row][13], 3));
                            sumWeightOS2 = (float)(sumWeightOS2 + Math.Round((float)dataTable.Rows[row][14]));
                            sumVolumeRej2 = (float)(sumVolumeRej2 + Math.Round((float)dataTable.Rows[row][15], 3));
                            sumWeightRej2 = (float)(sumWeightRej2 + Math.Round((float)dataTable.Rows[row][16]));

                            if ((bool)dataTable.Rows[row][17])
                            {
                                sumRegarding2 = (float)(sumRegarding2 + Math.Round((float)dataTable.Rows[row][11]));
                            }
                        }
                    }
                    else // 2 смена
                    {
                        if (shiftsDays.Count == 2)
                        {
                            sumPack2 += (int)dataTable.Rows[row][9];
                            sumVolume2 = (float)(sumVolume2 + Math.Round((float)dataTable.Rows[row][10], 3));
                            sumWeight2 = (float)(sumWeight2 + Math.Round((float)dataTable.Rows[row][11]));
                            sumOS2 += (int)dataTable.Rows[row][12];
                            sumVolumeOS2 = (float)(sumVolumeOS2 + Math.Round((float)dataTable.Rows[row][13], 3));
                            sumWeightOS2 = (float)(sumWeightOS2 + Math.Round((float)dataTable.Rows[row][14]));
                            sumVolumeRej2 = (float)(sumVolumeRej2 + Math.Round((float)dataTable.Rows[row][15], 3));
                            sumWeightRej2 = (float)(sumWeightRej2 + Math.Round((float)dataTable.Rows[row][16]));

                            if ((bool)dataTable.Rows[row][17])
                            {
                                sumRegarding2 = (float)(sumRegarding2 + Math.Round((float)dataTable.Rows[row][11]));
                            }
                        }
                    }
                }
                // блок подсчета сумм
                int[] columnsSheet1 = { 8, 10, 11, 13, 15, 17, 19, 20 }; // столбцы сумм смена
                int[] columnsSheet2 = { 11, 12, 13, 15, 16, 18, 19, 20 }; // столбцы сумм сутки
                double[] sums = { sumPack1, sumPack2, sumVolume1, sumVolume2, sumWeight1, sumWeight2, sumOS1, sumOS2, sumVolumeOS1, sumVolumeOS2,
            sumWeightOS1, sumWeightOS2, sumVolumeRej1, sumVolumeRej2, sumWeightRej1, sumWeightRej2 };

                for (int i = 0; i < columnsSheet1.Length; i++)
                {
                    worksheet.Cell(21, columnsSheet1[i]).Value = worksheet.Cell(97, columnsSheet2[i]).Value = (sums[i * 2]);
                    worksheet.Cell(51, columnsSheet1[i]).Value = worksheet.Cell(98, columnsSheet2[i]).Value = (sums[i * 2 + 1]);
                    worksheet.Cell(99, columnsSheet2[i]).Value = (sums[i * 2]) + (sums[i * 2 + 1]);
                }
                // общий итог смена 1
                double allWeight1 = sumWeight1 + sumWeightOS1 + sumWeightRej1;
                worksheet.Cell(22, 4).Value = allWeight1;
                if (allWeight1 != 0)
                {
                    worksheet.Cell(22, 12).Value = Math.Round(sumWeightOS1 / allWeight1 * 100, 2);
                    worksheet.Cell(22, 17).Value = Math.Round(sumWeightRej1 / allWeight1 * 100, 2);
                    worksheet.Cell(22, 20).Value = Math.Round(sumRegarding1 / allWeight1 * 100, 2);
                }
                else
                {
                    worksheet.Cell(22, 12).Value = 0;
                    worksheet.Cell(22, 17).Value = 0;
                    worksheet.Cell(22, 20).Value = 0;
                }
                // общий итог смена 2
                double allWeight2 = sumWeight2 + sumWeightOS2 + sumWeightRej2;
                if (allWeight2 != 0)
                {
                    worksheet.Cell(52, 4).Value = allWeight2;
                    worksheet.Cell(52, 12).Value = Math.Round(sumWeightOS2 / allWeight2 * 100, 2);
                    worksheet.Cell(52, 17).Value = Math.Round(sumWeightRej2 / allWeight2 * 100, 2);
                    worksheet.Cell(52, 20).Value = Math.Round(sumRegarding2 / allWeight2 * 100, 2);
                }
                else
                {
                    worksheet.Cell(52, 4).Value = 0;
                    worksheet.Cell(52, 12).Value = 0;
                    worksheet.Cell(52, 17).Value = 0;
                    worksheet.Cell(52, 20).Value = 0;
                }
                // общий итог сутки
                double allWeight3 = allWeight1 + allWeight2;
                worksheet.Cell(100, 4).Value = allWeight3;
                worksheet.Cell(100, 8).Value = sumVolume1 + sumVolumeOS1 + sumVolumeRej1 + sumVolume2 + sumVolumeOS2 + sumVolumeRej2;
                if (allWeight3 != 0)
                {
                    worksheet.Cell(100, 12).Value = Math.Round((sumWeightOS1 + sumWeightOS2) / allWeight3 * 100, 2);
                    worksheet.Cell(100, 17).Value = Math.Round((sumWeightRej1 + sumWeightRej2) / allWeight3 * 100, 2);
                    worksheet.Cell(100, 20).Value = Math.Round((sumRegarding1 + sumRegarding2) / allWeight3 * 100, 2);
                }
                else
                {
                    worksheet.Cell(100, 12).Value = 0;
                    worksheet.Cell(100, 17).Value = 0;
                    worksheet.Cell(100, 20).Value = 0;
                }
                 workbook.Save();
                
                ProjectLogger.LogDebug("Конец метода WriteToExcel24hours");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                ProjectLogger.LogException("Ошибка в методе WriteToExcel24hours", ex);
            }
        }
        //---------------------------------------------------------------------------------------------------------------------------------------
        public string GetProductNameDay(DataTable dataTable, int row) // Выведение имени продукта
        {

            string productName;
            if ((bool)dataTable.Rows[row][3]) // блок названия продукта
            {
                if (dataTable.Rows[row][2].ToString().Contains("М"))
                {
                    string lastThreeChars = dataTable.Rows[row][2].ToString().Split('М')[1];
                    productName = $"{dataTable.Rows[row][2].ToString().Split('М')[0]} {dataTable.Rows[row][5]}x{dataTable.Rows[row][6]}x{dataTable.Rows[row][4]} ({lastThreeChars})";
                }
                else
                {
                    productName = $"{dataTable.Rows[row][2]} {dataTable.Rows[row][5]}x{dataTable.Rows[row][6]}x{dataTable.Rows[row][4]}";
                }
            }
            else
            {
                productName = $"{dataTable.Rows[row][2]} {dataTable.Rows[row][5]}x{dataTable.Rows[row][6]}x{dataTable.Rows[row][4]}";
            }
            return productName;
        }
        //---------------------------------------------------------------------------------------------------------------------------------------
        public string GetProductName(DataTable dataTable, int row) // Выведение имени продукта
        {

            string productName;
            if ((bool)dataTable.Rows[row][1]) // блок названия продукта
            {
                if (dataTable.Rows[row][0].ToString().Contains("М"))
                {
                    string lastThreeChars = dataTable.Rows[row][0].ToString().Split('М')[1];
                    productName  = $"{dataTable.Rows[row][0].ToString().Split('М')[0]} {dataTable.Rows[row][3]}x{dataTable.Rows[row][4]}x{dataTable.Rows[row][2]} ({lastThreeChars})";
                }
                else
                {
                    productName = $"{dataTable.Rows[row][0]} {dataTable.Rows[row][3]}x{dataTable.Rows[row][4]}x{dataTable.Rows[row][2]}";
                }
            }
            else
            {
                productName = $"{dataTable.Rows[row][0]} {dataTable.Rows[row][3]}x{dataTable.Rows[row][4]}x{dataTable.Rows[row][2]}";
            }
            return productName ;
        }
        //---------------------------------------------------------------------------------------------------------------------------------------
        public void WriteToExcelInToDates(XLWorkbook workbook, DataTable dataTable, int getMethod, List<string> prodList, List<int> shiftsDays, List<int> shifts, List<string> stopCategoryes,
            DateTime varDate1 = default, DateTime varDate2 = default)
        {
            try
            {
                ProjectLogger.LogDebug("Начало метода WriteToExcelInToDates");
                var worksheet = workbook.Worksheet(1);
                int firstRow = 8, countRow = 0;
                int summaryRow = 49;

                int sumPack = 0, sumOS = 0;
                float sumVolume = 0, sumWeight = 0, sumVolumeOS = 0, sumWeightOS = 0;
                float sumVolumeRej = 0, sumWeightRej = 0;
                double allWeight = 0, allVolume = 0;

                int[] columnsSheet = { 11, 12, 13, 15, 16, 18, 19, 20 }; // столбцы сумм сутки
                double[] sums = { sumPack, sumVolume, sumWeight, sumOS, sumVolumeOS, sumWeightOS, sumVolumeRej, sumWeightRej };
                if (getMethod == 0)
                {

                    for (int row = 0; row < dataTable.Rows.Count; row++) // блок заполнения смен
                    {
                        if (prodList.Contains(GetProductName(dataTable, row)))
                        {
                            worksheet.Cell(firstRow - 3, 8).Value = varDate1;
                            worksheet.Cell(firstRow - 3, 11).Value = varDate2.AddDays(1);
                            if (countRow != 0 && countRow % 40 == 0)
                            {
                                firstRow += 24;
                                summaryRow += 54;
                                sumPack = 0; sumOS = 0;
                                sumVolume = 0; sumWeight = 0; sumVolumeOS = 0; sumWeightOS = 0;
                                sumVolumeRej = 0; sumWeightRej = 0;
                                allWeight = 0; allVolume = 0;
                            }
                            countRow++;
                            for (int item = 0; item > sums.Length; item++)
                            {
                                sums[item] = 0;
                            };

                            worksheet.Cell(countRow + firstRow, 1).Value = countRow; // порядковый номер продукта (в Кандыгаше №партии
                            worksheet.Cell(countRow + firstRow, 3).Value = GetProductName(dataTable, row); // наименование продукта
                            worksheet.Cell(countRow + firstRow, 10).Value = (float)dataTable.Rows[row][5]; // Объем 1 пачки
                            worksheet.Cell(countRow + firstRow, 4).Value = (float)dataTable.Rows[row][6]; // средняя плотность
                            worksheet.Cell(countRow + firstRow, 11).Value = (int)dataTable.Rows[row][7]; // пачек гп
                            worksheet.Cell(countRow + firstRow, 12).Value = Math.Round((float)dataTable.Rows[row][8], 3); // объем
                            worksheet.Cell(countRow + firstRow, 13).Value = Math.Round((float)dataTable.Rows[row][9]); // вес
                            worksheet.Cell(countRow + firstRow, 6).Value = (float)dataTable.Rows[row][10]; // плотность ос
                            worksheet.Cell(countRow + firstRow, 15).Value = (int)dataTable.Rows[row][11]; // пачек ос
                            worksheet.Cell(countRow + firstRow, 16).Value = Math.Round((float)dataTable.Rows[row][12], 3); // Объем ос
                            worksheet.Cell(countRow + firstRow, 18).Value = Math.Round((float)dataTable.Rows[row][13]); // вес ос
                            worksheet.Cell(countRow + firstRow, 8).Value = (float)dataTable.Rows[row][10]; // плотность обрези
                            worksheet.Cell(countRow + firstRow, 19).Value = Math.Round((float)dataTable.Rows[row][14], 3); // Объем обрези
                            worksheet.Cell(countRow + firstRow, 20).Value = Math.Round((float)dataTable.Rows[row][15]); // вес обрези
                            worksheet.Cell(countRow + firstRow, 21).Value = Math.Round((double)dataTable.Rows[row][16], 2); // % OS
                            worksheet.Cell(countRow + firstRow, 22).Value = Math.Round((double)dataTable.Rows[row][17], 2); // % обрези

                            sums[0] = sumPack += (int)dataTable.Rows[row][7];
                            sums[1] = sumVolume = (float)(sumVolume + Math.Round((float)dataTable.Rows[row][8], 3));
                            sums[2] = sumWeight = (float)(sumWeight + Math.Round((float)dataTable.Rows[row][9]));
                            sums[3] = sumOS += (int)dataTable.Rows[row][11];
                            sums[4] = sumVolumeOS = (float)(sumVolumeOS + Math.Round((float)dataTable.Rows[row][12], 3));
                            sums[5] = sumWeightOS = (float)(sumWeightOS + Math.Round((float)dataTable.Rows[row][13]));
                            sums[6] = sumVolumeRej = (float)(sumVolumeRej + Math.Round((float)dataTable.Rows[row][14], 3));
                            sums[7] = sumWeightRej = (float)(sumWeightRej + Math.Round((float)dataTable.Rows[row][15]));

                        }
                        for (int i = 0; i < columnsSheet.Length; i++)
                        {
                            worksheet.Cell(summaryRow, columnsSheet[i]).Value = sums[i];
                        }
                        allWeight = sumWeight + sumWeightOS + sumWeightRej;
                        allVolume = sumVolume + sumVolumeOS + sumVolumeRej;
                        worksheet.Cell(summaryRow + 1, 4).Value = allWeight;
                        worksheet.Cell(summaryRow + 1, 8).Value = allVolume;
                        if (allWeight != 0)
                        {
                            worksheet.Cell(summaryRow + 1, 12).Value = Math.Round(sumWeightOS / allWeight * 100, 2);
                            worksheet.Cell(summaryRow + 1, 17).Value = Math.Round(sumWeightRej / allWeight * 100, 2);
                        }
                        else
                        {
                            worksheet.Cell(summaryRow + 1, 12).Value = 0;
                            worksheet.Cell(summaryRow + 1, 17).Value = 0;
                        }
                    }

                }
                if (getMethod == 3)
                {

                    int shiftNumConst = (int)dataTable.Rows[0][5];
                    for (int row = 0; row < dataTable.Rows.Count; row++) // блок заполнения смен
                    {
                        if (prodList.Contains(GetProductName(dataTable, row)))
                        {
                            worksheet.Cell(firstRow - 3, 8).Value = varDate1;
                            worksheet.Cell(firstRow - 3, 11).Value = varDate2.AddDays(1);

                            if (shiftNumConst != (int)dataTable.Rows[row][5])
                            {
                                shiftNumConst = (int)dataTable.Rows[row][5];
                                firstRow += 54;
                                summaryRow += 54;
                                countRow = 0;
                                sumPack = 0; sumOS = 0;
                                sumVolume = 0; sumWeight = 0; sumVolumeOS = 0; sumWeightOS = 0;
                                sumVolumeRej = 0; sumWeightRej = 0;
                                allWeight = 0; allVolume = 0;
                            }
                            countRow++;
                            for (int item = 0; item > sums.Length; item++)
                            {
                                sums[item] = 0;
                            };

                            worksheet.Cell(countRow + firstRow, 1).Value = countRow; // порядковый номер продукта (в Кандыгаше №партии
                            worksheet.Cell(countRow + firstRow, 2).Value = (int)dataTable.Rows[row][5];
                            worksheet.Cell(countRow + firstRow, 3).Value = GetProductName(dataTable, row); // наименование продукта
                            worksheet.Cell(countRow + firstRow, 10).Value = (float)dataTable.Rows[row][6]; // Объем 1 пачки
                            worksheet.Cell(countRow + firstRow, 4).Value = (float)dataTable.Rows[row][7]; // средняя плотность
                            worksheet.Cell(countRow + firstRow, 11).Value = (int)dataTable.Rows[row][8]; // пачек гп
                            worksheet.Cell(countRow + firstRow, 12).Value = Math.Round((float)dataTable.Rows[row][9], 3); // объем
                            worksheet.Cell(countRow + firstRow, 13).Value = Math.Round((float)dataTable.Rows[row][10]); // вес
                            worksheet.Cell(countRow + firstRow, 6).Value = (float)dataTable.Rows[row][11]; // плотность ос
                            worksheet.Cell(countRow + firstRow, 15).Value = (int)dataTable.Rows[row][12]; // пачек ос
                            worksheet.Cell(countRow + firstRow, 16).Value = Math.Round((float)dataTable.Rows[row][13], 3); // Объем ос
                            worksheet.Cell(countRow + firstRow, 18).Value = Math.Round((float)dataTable.Rows[row][14]); // вес ос
                            worksheet.Cell(countRow + firstRow, 8).Value = (float)dataTable.Rows[row][11]; // плотность обрези
                            worksheet.Cell(countRow + firstRow, 19).Value = Math.Round((float)dataTable.Rows[row][15], 3); // Объем обрези
                            worksheet.Cell(countRow + firstRow, 20).Value = Math.Round((float)dataTable.Rows[row][16]); // вес обрези
                            worksheet.Cell(countRow + firstRow, 21).Value = Math.Round((double)dataTable.Rows[row][17], 2); // % OS
                            worksheet.Cell(countRow + firstRow, 22).Value = Math.Round((double)dataTable.Rows[row][18], 2); // % обрези

                            sums[0] = sumPack += (int)dataTable.Rows[row][8];
                            sums[1] = sumVolume = (float)(sumVolume + Math.Round((float)dataTable.Rows[row][9], 3));
                            sums[2] = sumWeight = (float)(sumWeight + Math.Round((float)dataTable.Rows[row][10]));
                            sums[3] = sumOS += (int)dataTable.Rows[row][12];
                            sums[4] = sumVolumeOS = (float)(sumVolumeOS + Math.Round((float)dataTable.Rows[row][13], 3));
                            sums[5] = sumWeightOS = (float)(sumWeightOS + Math.Round((float)dataTable.Rows[row][14]));
                            sums[6] = sumVolumeRej = (float)(sumVolumeRej + Math.Round((float)dataTable.Rows[row][15], 3));
                            sums[7] = sumWeightRej = (float)(sumWeightRej + Math.Round((float)dataTable.Rows[row][16]));

                        }
                        for (int i = 0; i < columnsSheet.Length; i++)
                        {
                            worksheet.Cell(summaryRow, columnsSheet[i]).Value = sums[i];
                        }
                        allWeight = sumWeight + sumWeightOS + sumWeightRej;
                        allVolume = sumVolume + sumVolumeOS + sumVolumeRej;
                        worksheet.Cell(summaryRow + 1, 4).Value = allWeight;
                        worksheet.Cell(summaryRow + 1, 8).Value = allVolume;
                        if (allWeight != 0)
                        {
                            worksheet.Cell(summaryRow + 1, 12).Value = Math.Round(sumWeightOS / allWeight * 100, 2);
                            worksheet.Cell(summaryRow + 1, 17).Value = Math.Round(sumWeightRej / allWeight * 100, 2);
                        }
                        else
                        {
                            worksheet.Cell(summaryRow + 1, 12).Value = 0;
                            worksheet.Cell(summaryRow + 1, 17).Value = 0;
                        }
                    }
                }
                workbook.Save();
                ProjectLogger.LogDebug("Конец метода WriteToExcelInToDates");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                ProjectLogger.LogException("Ошибка в методе WriteToExcelInToDates", ex);
            }
        }
        //---------------------------------------------------------------------------------------------------------------------------------------
        public void WriteToExcelBrokenProduct(XLWorkbook workbook, DataTable dataTable, int getMethod, List<int> shiftsDays, List<int> shifts, List<string> stopCategoryes,
            DateTime varDate1 = default, DateTime varDate2 = default)
        {
            try
            {
                ProjectLogger.LogDebug("Начало метода WriteToExcelBrokenProduct");
                var worksheet = workbook.Worksheet(1);
                int firstRow = 8, countRow = 0;
                int summaryRow = 21;

                int sumPack = 0;
                float sumVolume = 0, sumWeight = 0, sumWeightOS = 0;

                int[] columnsSheet = { 11, 14, 18 }; // столбцы сумм сутки
                double[] sums = { sumPack, sumVolume, sumWeight };
                if (getMethod == 4)
                {
                    for (int row = 0; row < dataTable.Rows.Count; row++) // блок заполнения смен
                    {
                        worksheet.Cell(firstRow - 3, 8).Value = varDate1;
                        worksheet.Cell(firstRow - 3, 11).Value = varDate2.AddDays(1);
                        countRow++;

                        worksheet.Cell(countRow + firstRow, 1).Value = (int)dataTable.Rows[row][0]; // код дефекта
                        worksheet.Cell(countRow + firstRow, 2).Value = (string)dataTable.Rows[row][1]; // Наименование дефекта
                        worksheet.Cell(countRow + firstRow, 11).Value = (int)dataTable.Rows[row][2]; // пачек ос
                        worksheet.Cell(countRow + firstRow, 14).Value = Math.Round((float)dataTable.Rows[row][3], 3); // Объем ос
                        worksheet.Cell(countRow + firstRow, 18).Value = Math.Round((float)dataTable.Rows[row][4]); // вес ос

                        sums[0] = sumPack += (int)dataTable.Rows[row][2];
                        sums[1] = sumVolume = (float)(sumVolume + Math.Round((float)dataTable.Rows[row][3], 3));
                        sums[2] = sumWeight = (float)(sumWeight + Math.Round((float)dataTable.Rows[row][4]));
                        if ((int)dataTable.Rows[row][0] < 8)
                        {
                            sumWeightOS = (float)(sumWeightOS + Math.Round((float)dataTable.Rows[row][4]));
                        }

                        for (int i = 0; i < columnsSheet.Length; i++)
                        {
                            worksheet.Cell(summaryRow, columnsSheet[i]).Value = sums[i];
                        }
                    }
                    for (int row1 = 0; row1 < dataTable.Rows.Count; row1++) // блок заполнения смен
                    {
                        worksheet.Cell(row1 + 1 + firstRow, 9).Value = sumWeightOS != 0 ? Math.Round((float)dataTable.Rows[row1][4]) / sumWeightOS : 0;
                    }
                }
                if (getMethod == 5)
                {

                    int shiftNumConst = (int)dataTable.Rows[0][0];
                    for (int row = 0; row < dataTable.Rows.Count; row++) // блок заполнения
                    {
                        if (shiftNumConst != (int)dataTable.Rows[row][0])
                        {
                            PercentBrokenCategory(workbook, dataTable, firstRow, sumWeightOS, shiftNumConst);
                            shiftNumConst = (int)dataTable.Rows[row][0];
                            firstRow += 27;
                            summaryRow += 27;
                            countRow = 0;
                            sumPack = 0;
                            sumVolume = 0; sumWeight = 0; sumWeightOS = 0;
                        }
                        countRow++;
                        for (int item = 0; item > sums.Length; item++)
                        {
                            sums[item] = 0;
                        };

                        worksheet.Cell(firstRow - 3, 4).Value = (int)dataTable.Rows[row][0];
                        worksheet.Cell(firstRow - 3, 8).Value = varDate1;
                        worksheet.Cell(firstRow - 3, 11).Value = varDate2.AddDays(1);

                        worksheet.Cell(countRow + firstRow, 1).Value = (int)dataTable.Rows[row][1]; ; // код дефекта
                        worksheet.Cell(countRow + firstRow, 2).Value = (string)dataTable.Rows[row][2]; // Наименование дефекта
                        worksheet.Cell(countRow + firstRow, 11).Value = (int)dataTable.Rows[row][3]; // пачек ос
                        worksheet.Cell(countRow + firstRow, 14).Value = Math.Round((float)dataTable.Rows[row][4], 3); // Объем ос
                        worksheet.Cell(countRow + firstRow, 18).Value = Math.Round((float)dataTable.Rows[row][5]); // вес ос

                        sums[0] = sumPack += (int)dataTable.Rows[row][3];
                        sums[1] = sumVolume = (float)(sumVolume + Math.Round((float)dataTable.Rows[row][4], 3));
                        sums[2] = sumWeight = (float)(sumWeight + Math.Round((float)dataTable.Rows[row][5]));
                        if ((int)dataTable.Rows[row][0] < 8)
                        {
                            sumWeightOS = (float)(sumWeightOS + Math.Round((float)dataTable.Rows[row][5]));
                        }

                        for (int i = 0; i < columnsSheet.Length; i++)
                        {
                            worksheet.Cell(summaryRow, columnsSheet[i]).Value = sums[i];
                        }
                    }
                    PercentBrokenCategory(workbook, dataTable, firstRow, sumWeightOS, shiftNumConst);
                }
                workbook.Save();
                ProjectLogger.LogDebug("Конец метода WriteToExcelBrokenProduct");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                ProjectLogger.LogException("Ошибка в методе WriteToExcelBrokenProduct", ex);
            }
        }
        //---------------------------------------------------------------------------------------------------------------------------------------
        public void PercentBrokenCategory(XLWorkbook workbook, DataTable dataTable, int firstRow, float sumWeightOS, int shiftNumStat) // Выведение имени продукта
        {
            var worksheet = workbook.Worksheet(1);
            int count = 0;
            for(int row = 0; row < dataTable.Rows.Count; row++)
            {
                if ((int)dataTable.Rows[row][0] == shiftNumStat)
                {
                    count++;
                    worksheet.Cell(count + firstRow, 9).Value = sumWeightOS != 0 ? Math.Round((float)dataTable.Rows[row][5]) / sumWeightOS : 0 ;
                }
            }
        }
        //---------------------------------------------------------------------------------------------------------------------------------------
        public void WriteToExcelShiftStops(XLWorkbook workbook, DataTable dataTable, int getMethod, List<int> shiftsDays, List<int> shifts, List<string> stopCategoryes,
             DateTime varDate1 = default, DateTime varDate2 = default)
        {
            try
            {
                ProjectLogger.LogDebug("Начало метода WriteToExcel24hours");
                var worksheet = workbook.Worksheet(1);
                int firstRow2 = 0, countRow1 = 0;
                int shiftNumConst = (int)dataTable.Rows[0][1];

                DateTime shiftDate = varDate1.Date;
                shiftDate.ToShortDateString();
                worksheet.Cell(5, 3).Value = shiftDate;
                worksheet.Cell(5, 9).Value = shiftNumConst;
                worksheet.Cell(5, 5).Value = shiftsDays[0] == 1 ? "8:00-20:00" : "20:00-8:00";

                float sumDurationStopProd = 0, sumDurationRunProd = 0;
                for (int row = 0; row < dataTable.Rows.Count; row++) // блок заполнения смен
                {
                    countRow1++;
                    firstRow2 = 7;

                    worksheet.Cell(countRow1 + firstRow2, 1).Value = countRow1; // порядковый номер продукта (в Кандыгаше №партии
                    worksheet.Cell(countRow1 + firstRow2, 2).Value = (int)dataTable.Rows[row][0]; // смена
                    worksheet.Cell(countRow1 + firstRow2, 3).Value = (string)dataTable.Rows[row][1]; // место
                    worksheet.Cell(countRow1 + firstRow2, 4).Value = (string)dataTable.Rows[row][2]; // узел
                    worksheet.Cell(countRow1 + firstRow2, 5).Value = (string)dataTable.Rows[row][3]; // тип простоя
                    worksheet.Cell(countRow1 + firstRow2, 9).Value = (string)dataTable.Rows[row][4]; // остановка
                    worksheet.Cell(countRow1 + firstRow2, 11).Value = (string)dataTable.Rows[row][5]; // начало остановки
                    worksheet.Cell(countRow1 + firstRow2, 12).Value = (string)dataTable.Rows[row][6]; // конец остановки
                    worksheet.Cell(countRow1 + firstRow2, 13).Value = (float)dataTable.Rows[row][7]; // длительность остановки
                    worksheet.Cell(countRow1 + firstRow2, 14).Value = (bool)dataTable.Rows[row][8] == true ? "нет" : "да"; // влияние остановки
                    worksheet.Cell(countRow1 + firstRow2, 15).Value = (string)dataTable.Rows[row][9]; // комментарий
                    if ((bool)dataTable.Rows[row][8] != true) sumDurationStopProd += (float)dataTable.Rows[row][7];
                    else sumDurationRunProd += (float)dataTable.Rows[row][7];
                }
                float sumDuration = sumDurationStopProd + sumDurationRunProd;

                worksheet.Cell(51, 13).Value = worksheet.Cell(8, 7).Value = sumDurationStopProd;
                worksheet.Cell(51, 14).Value = sumDurationRunProd;
                worksheet.Cell(52, 13).Value = sumDuration;
                worksheet.Cell(8, 1).Value = 720 - sumDurationStopProd;

                workbook.Save();

                ProjectLogger.LogDebug("Конец метода WriteToExcelPivot");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                ProjectLogger.LogException("Ошибка в методе WriteToExcelPivot", ex);
            }
        }
        //---------------------------------------------------------------------------------------------------------------------------------------
        public void WriteToExcelPivote(XLWorkbook workbook, List<DataTable> dataTablePivot, int getMethod, List<int> shiftsDays, List<int> shifts, List<string> stopCategoryes,
             DateTime varDate1 = default, DateTime varDate2 = default)
        {
            try
            {
                ProjectLogger.LogDebug("Начало метода WriteToExcel24hours");
                var worksheet = workbook.Worksheet(1);
                int firstRow = 0, firstRow2 = 0, countRow = 0, countRow1 = 0;
                DataTable dataTable1 = dataTablePivot[0];
                DataTable dataTable2 = dataTablePivot[1];
                int shiftNumConst = (int)dataTable1.Rows[0][1];

                DateTime shiftDate = varDate1.Date;
                shiftDate.ToShortDateString();
                worksheet.Cell(5, 3).Value = shiftDate;
                worksheet.Cell(5, 9).Value = shiftNumConst;
                worksheet.Cell(5, 5).Value = shiftsDays[0] == 1 ? "8:00-20:00" : "20:00-8:00";

                int sumPack1 = 0, sumOS1 = 0;
                float sumVolume1 = 0, sumWeight1 = 0, sumVolumeOS1 = 0, sumWeightOS1 = 0;
                float sumVolumeRej1 = 0, sumWeightRej1 = 0, sumRegarding1 = 0;

                for (int row = 0; row < dataTable1.Rows.Count; row++) // блок заполнения смен
                {
                    countRow++;
                    if ((bool)dataTable1.Rows[row][17]) // подсчет переорта см 1
                    {
                        sumRegarding1 = (float)(sumRegarding1 + Math.Round((float)dataTable1.Rows[row][11]));
                    }
                    
                    firstRow = 19; // выбор первой строки сменных от смены

                    worksheet.Cell(countRow + firstRow, 1).Value = countRow; // порядковый номер продукта (в Кандыгаше №партии
                    worksheet.Cell(countRow + firstRow, 2).Value = (int)dataTable1.Rows[row][1]; // смена
                    worksheet.Cell(countRow + firstRow, 3).Value = GetProductNameDay(dataTable1, row); // наименование продукта
                    worksheet.Cell(countRow + firstRow, 10).Value = (float)dataTable1.Rows[row][8]; // Объем 1 пачки
                    worksheet.Cell(countRow + firstRow, 4).Value = (float)dataTable1.Rows[row][7]; // средняя плотность
                    worksheet.Cell(countRow + firstRow, 11).Value = (int)dataTable1.Rows[row][9]; // пачек гп
                    worksheet.Cell(countRow + firstRow, 12).Value = Math.Round((float)dataTable1.Rows[row][10], 3); // объем
                    worksheet.Cell(countRow + firstRow, 13).Value = Math.Round((float)dataTable1.Rows[row][11]); // вес
                    worksheet.Cell(countRow + firstRow, 6).Value = (float)dataTable1.Rows[row][18]; // плотность ос
                    worksheet.Cell(countRow + firstRow, 15).Value = (int)dataTable1.Rows[row][12]; // пачек ос
                    worksheet.Cell(countRow + firstRow, 16).Value = Math.Round((float)dataTable1.Rows[row][13], 3); // Объем ос
                    worksheet.Cell(countRow + firstRow, 18).Value = Math.Round((float)dataTable1.Rows[row][14]); // вес ос
                    worksheet.Cell(countRow + firstRow, 8).Value = (float)dataTable1.Rows[row][18]; // плотность обрези
                    worksheet.Cell(countRow + firstRow, 19).Value = Math.Round((float)dataTable1.Rows[row][15], 3); // Объем обрези
                    worksheet.Cell(countRow + firstRow, 20).Value = Math.Round((float)dataTable1.Rows[row][16]); // вес обрези

                    if (shiftNumConst == (int)dataTable1.Rows[row][1]) // суммирование
                    {
                        sumPack1 += (int)dataTable1.Rows[row][9];
                        sumVolume1 = (float)(sumVolume1 + Math.Round((float)dataTable1.Rows[row][10], 3));
                        sumWeight1 = (float)(sumWeight1 + Math.Round((float)dataTable1.Rows[row][11]));
                        sumOS1 += (int)dataTable1.Rows[row][12];
                        sumVolumeOS1 = (float)(sumVolumeOS1 + Math.Round((float)dataTable1.Rows[row][13], 3));
                        sumWeightOS1 = (float)(sumWeightOS1 + Math.Round((float)dataTable1.Rows[row][14]));
                        sumVolumeRej1 = (float)(sumVolumeRej1 + Math.Round((float)dataTable1.Rows[row][15], 3));
                        sumWeightRej1 = (float)(sumWeightRej1 + Math.Round((float)dataTable1.Rows[row][16]));

                        if ((bool)dataTable1.Rows[row][17])
                        {
                            sumRegarding1 = (float)(sumRegarding1 + Math.Round((float)dataTable1.Rows[row][11]));
                        }
                    }
                }
                // блок подсчета сумм
                int[] columnsSheet2 = { 11, 12, 13, 15, 16, 18, 19, 20 }; // столбцы сумм сутки
                double[] sums = { sumPack1, sumVolume1, sumWeight1, sumOS1, sumVolumeOS1, sumWeightOS1, sumVolumeRej1, sumWeightRej1};

                for (int i = 0; i < columnsSheet2.Length; i++)
                {
                    worksheet.Cell(35, columnsSheet2[i]).Value = sums[i];
                }
                // общий итог смена 1
                double allWeight1 = sumWeight1 + sumWeightOS1 + sumWeightRej1;
                double allVolume1 = sumVolume1 + sumVolumeOS1 + sumVolumeRej1;
                worksheet.Cell(36, 4).Value = allWeight1;
                worksheet.Cell(36, 8).Value = allVolume1;
                if (allWeight1 != 0)
                {
                    worksheet.Cell(36, 12).Value = Math.Round(sumWeightOS1 / allWeight1 * 100, 2);
                    worksheet.Cell(36, 16).Value = Math.Round(sumWeightRej1 / allWeight1 * 100, 2);
                    worksheet.Cell(36, 20).Value = Math.Round(sumRegarding1 / allWeight1 * 100, 2);
                }
                else
                {
                    worksheet.Cell(36, 12).Value = 0;
                    worksheet.Cell(36, 17).Value = 0;
                    worksheet.Cell(36, 20).Value = 0;
                }

                float sumDurationStopProd = 0, sumDurationRunProd = 0, numFuge1 = 0, numFuge2 = 0;
                string whyChangeFuge1 = "", whyChangeFuge2 = "";
                string timeFuge1 = "", timeFuge2 = "";
                for (int row = 0; row < dataTable2.Rows.Count; row++) // блок заполнения смен
                {
                    countRow1++;
                    firstRow2 = 38;

                    worksheet.Cell(countRow1 + firstRow2, 1).Value = countRow1; // порядковый номер продукта (в Кандыгаше №партии
                    worksheet.Cell(countRow1 + firstRow2, 2).Value = (int)dataTable2.Rows[row][0]; // смена
                    worksheet.Cell(countRow1 + firstRow2, 3).Value = (string)dataTable2.Rows[row][1]; // место
                    worksheet.Cell(countRow1 + firstRow2, 4).Value = (string)dataTable2.Rows[row][2]; // узел
                    worksheet.Cell(countRow1 + firstRow2, 5).Value = (string)dataTable2.Rows[row][3]; // тип простоя
                    worksheet.Cell(countRow1 + firstRow2, 9).Value = (string)dataTable2.Rows[row][4]; // остановка
                    worksheet.Cell(countRow1 + firstRow2, 11).Value = (string)dataTable2.Rows[row][5]; // начало остановки
                    worksheet.Cell(countRow1 + firstRow2, 12).Value = (string)dataTable2.Rows[row][6]; // конец остановки
                    worksheet.Cell(countRow1 + firstRow2, 13).Value = (float)dataTable2.Rows[row][7]; // длительность остановки
                    worksheet.Cell(countRow1 + firstRow2, 14).Value = (bool)dataTable2.Rows[row][8] == true ? "нет" : "да"; // влияние остановки
                    worksheet.Cell(countRow1 + firstRow2, 15).Value = (string)dataTable2.Rows[row][9]; // комментарий
                    if ((bool)dataTable2.Rows[row][8] != true) sumDurationStopProd += (float)dataTable2.Rows[row][7];
                    else sumDurationRunProd += (float)dataTable2.Rows[row][7];
                    if ((int)dataTable2.Rows[row][11] > 0 && numFuge1 == 0)
                    {
                        whyChangeFuge1 = (string)dataTable2.Rows[row][9];
                        numFuge1 = (int)dataTable2.Rows[row][11];
                        timeFuge1 = (string)dataTable2.Rows[row][10];
                    }
                    else if ((int)dataTable2.Rows[row][11] > 0 && numFuge1 != 0)
                    {
                        whyChangeFuge2 = (string)dataTable2.Rows[row][9];
                        numFuge2 = (int)dataTable2.Rows[row][11];
                        timeFuge2 = (string)dataTable2.Rows[row][10];
                    }
                }
                float sumDuration = sumDurationStopProd + sumDurationRunProd;

                worksheet.Cell(51, 13).Value = worksheet.Cell(8, 7).Value = sumDurationStopProd;
                worksheet.Cell(51, 14).Value = sumDurationRunProd;
                worksheet.Cell(52, 13).Value = sumDuration;
                worksheet.Cell(8, 1).Value = 720 - sumDurationStopProd;
                worksheet.Cell(8, 14).Value = Math.Round(allWeight1 / (720 - sumDurationStopProd)*60);
                if (numFuge1 > 0)
                {
                    worksheet.Cell(11, 4).Value = timeFuge1.ToString().Split(' ')[1];
                    worksheet.Cell(11, 7).Value = numFuge1;
                    worksheet.Cell(11, 8).Value = whyChangeFuge1;
                }
                if (numFuge2 > 0)
                {
                    worksheet.Cell(12, 4).Value = timeFuge2.ToString().Split(' ')[1];
                    worksheet.Cell(12, 7).Value = numFuge2;
                    worksheet.Cell(11, 14).Value = whyChangeFuge2;
                }

                workbook.Save();

                ProjectLogger.LogDebug("Конец метода WriteToExcelPivot");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                ProjectLogger.LogException("Ошибка в методе WriteToExcelPivot", ex);
            }
        }
    }
}