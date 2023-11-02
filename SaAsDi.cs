using System;
using Npgsql;
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
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Office.Word;
using DocumentFormat.OpenXml.Spreadsheet;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.StartPanel;
using DocumentFormat.OpenXml.Drawing;
using Path = System.IO.Path;
using DocumentFormat.OpenXml.Bibliography;

namespace ShiftReportApp1
{
    internal class SaAsDi
    {
        public DateTime VarDate1 { get; set; }
        public DateTime VarDate2 { get; set; }
        public void SaveExcelFile(int numbQuery, List<string> prodList, DateTime varDate1 = default, DateTime varDate2 = default,
            int numShift1 = 0, int numShift2 = 0, int numShift3 = 0, int numShift4 = 0,
             string typeStops1 = "_", string typeStops2 = "_", string typeStops3 = "_", string typeStops4 = "_")
        {
            VarDate1 = varDate1;
            VarDate2 = varDate2;

            try
            {
                ProjectLogger.LogDebug("Начало SaveExcelFile");
                string query = Names.Request(numbQuery);
                using (SaveFileDialog saveFileDialog = new SaveFileDialog())
                {
                    saveFileDialog.Filter = "Файлы Excel (*.xlsx)|*.xlsx|Все файлы (*.*)|*.*";
                    saveFileDialog.FilterIndex = 1;
                    saveFileDialog.RestoreDirectory = true;

                    if (saveFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        string filePath = saveFileDialog.FileName;
                        string templateFilePath = "";
                        // Копируем файл шаблона в папку с исполняемым файлом
                        if (numbQuery == 6 || numbQuery == 11)
                        {
                            templateFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "qtemp2.xlsx");
                        }
                        else if (numbQuery == 7 || numbQuery == 8 || numbQuery == 9 || numbQuery == 10)
                        {
                            templateFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "qtemp1.xlsx");
                        }
                        else if (numbQuery == 12 || numbQuery == 13)
                        {
                            templateFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "qtemp3.xlsx");
                        }
                        else if (numbQuery >= 21 && numbQuery <= 29)
                        {
                            templateFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "qtemp4.xlsx");
                        }


                        File.Copy(templateFilePath, Path.Combine(Path.GetDirectoryName(filePath), Path.GetFileNameWithoutExtension(filePath) + ".xlsx"), true);

                        using (var workbook = new XLWorkbook(Path.Combine(Path.GetDirectoryName(filePath), Path.GetFileNameWithoutExtension(filePath) + ".xlsx")))
                        {
                            DataBaseConnection dbConnection = new DataBaseConnection();
                            NpgsqlConnection connection = dbConnection.GetConnection();

                            try
                            {
                                connection.Open();
                                using (NpgsqlCommand cmd = new NpgsqlCommand(query, connection))
                                {
                                    cmd.Parameters.AddWithValue("@varDate1", varDate1);
                                    cmd.Parameters.AddWithValue("@varDate2", varDate2);
                                    cmd.Parameters.AddWithValue("@numShift1", numShift1);
                                    cmd.Parameters.AddWithValue("@numShift2", numShift2);
                                    cmd.Parameters.AddWithValue("@numShift3", numShift3);
                                    cmd.Parameters.AddWithValue("@numShift4", numShift4);

                                    cmd.Parameters.AddWithValue("@typeStops1", typeStops1);
                                    cmd.Parameters.AddWithValue("@typeStops2", typeStops2);
                                    cmd.Parameters.AddWithValue("@typeStops3", typeStops3);
                                    cmd.Parameters.AddWithValue("@typeStops4", typeStops4);

                                    if (numbQuery == 6)
                                    {
                                        int firstRow = 8;
                                        SaveVariance1(workbook, cmd, prodList, firstRow);
                                    }
                                    if (numbQuery == 7)
                                    {
                                        int firstRow = 6;
                                        SaveVariance2(workbook, cmd, prodList, firstRow);
                                    }
                                    if (numbQuery == 11)
                                    {
                                        int firstRow = 8;
                                        SaveVariance5(workbook, cmd, prodList, firstRow);
                                    }
                                    if (numbQuery == 8)
                                    {
                                        int firstRow = 35;
                                        SaveVariance4(workbook, cmd, prodList, firstRow);
                                    }
                                    if (numbQuery == 10 || numbQuery == 9)
                                    {
                                        int firstRow = 6;
                                        SaveVariance4(workbook, cmd, prodList, firstRow);
                                    }
                                    if (numbQuery == 12)
                                    {
                                        int firstRow = 8;
                                        SaveVariance3(workbook, cmd, prodList, firstRow);
                                    }
                                    if (numbQuery == 13)
                                    {
                                        int firstRow = 8;
                                        SaveVariance7(workbook, cmd, prodList, firstRow);
                                    }
                                    if (numbQuery == 21)
                                    {
                                        int firstRow = 8;
                                        SaveVariance6(workbook, cmd, firstRow);
                                    }
                                    if (numbQuery == 22)
                                    {
                                        int firstRow = 8;
                                        SaveVariance8(workbook, cmd, firstRow);
                                    }
                                    if (numbQuery == 23)
                                    {
                                        int firstRow = 8;
                                        SaveVariance9(workbook, cmd, firstRow);
                                    }
                                    if (numbQuery == 24)
                                    {
                                        int firstRow = 8;
                                        SaveVariance10(workbook, cmd, firstRow);
                                    }
                                    if (numbQuery == 25)
                                    {
                                        int firstRow = 8;
                                        SaveVariance11(workbook, cmd, firstRow);
                                    }
                                    if (numbQuery == 26)
                                    {
                                        int firstRow = 8;
                                        SaveVariance12(workbook, cmd, firstRow);
                                    }
                                    if (numbQuery == 27)
                                    {
                                        int firstRow = 8;
                                        SaveVariance12(workbook, cmd, firstRow);
                                    }
                                    if (numbQuery == 28)
                                    {
                                        int firstRow = 8;
                                        SaveVariance13(workbook, cmd, firstRow);
                                    }
                                    if (numbQuery == 29)
                                    {
                                        int firstRow = 8;
                                        SaveVariance13(workbook, cmd, firstRow);
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show($"An error occurred: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                        }

                        MessageBox.Show("Файл сохранен успешно.", "Успешно", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                ProjectLogger.LogDebug("Конец SaveExcelFile");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                ProjectLogger.LogException("Ошибка в SaveExcelFile", ex);
            }
        }

        //  -----------------------------------------------------------------------------------------
        // Чтение запроса. Выгрузка по продуктам за период
        private void SaveVariance1(XLWorkbook workbook, NpgsqlCommand cmd, List<string> prodList, int firstRow)
        {
            try
            {
                ProjectLogger.LogDebug("Начало чтения запроса SaveVariance1");
                using (NpgsqlDataReader reader = cmd.ExecuteReader())
                {
                    int iter = 0;
                    MethodReadProd methodRead = new MethodReadProd(); // Создаем экземпляр класса MethodRead

                    while (reader.Read())
                    {
                        // Заменяем блок инициализации переменных вызовом метода ProcessData
                        methodRead.ProcessData(reader, true, false, false); // Передаем reader и boolOption в метод

                        string outputString;

                        if (methodRead.Unspecified)
                        {
                            // Последние три символа productName
                            string lastThreeChars = methodRead.ProductName.Substring(methodRead.ProductName.Length - 3);
                            outputString = $"{methodRead.ProductName} {methodRead.Length}x{methodRead.Width}x{methodRead.Depth}({lastThreeChars})";
                        }
                        else
                        {
                            outputString = $"{methodRead.ProductName} {methodRead.Length}x{methodRead.Width}x{methodRead.Depth}";
                        }
                        if (prodList.Contains(outputString))
                        {
                            iter++;
                            // Теперь передаем methodRead вместо отдельных переменных в метод для сохранения в Excel
                            FillExcel filler = new FillExcel();
                            filler.WriteToExcel1(workbook, outputString, methodRead.AvgDensity, methodRead.CountGoodProd, methodRead.GoodProdVolume, methodRead.GoodProdWeight,
                                methodRead.CountLowQual, methodRead.LowQualVolume, methodRead.LowQualWeight, methodRead.RejectVolume, methodRead.RejectWeight,
                                methodRead.LowQualPercent, methodRead.RejectPercent, firstRow, iter, VarDate1, VarDate2);

                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                ProjectLogger.LogException("Ошибка в Чтение запроса. Выгрузка по продуктам за период", ex);
            }

        }

        //  -----------------------------------------------------------------------------------------
        // Чтение запроса. Выгрузка по продуктам за сутки
        private void SaveVariance2(XLWorkbook workbook, NpgsqlCommand cmd, List<string> prodList, int firstRow)
        {
            try
            {
                ProjectLogger.LogDebug("Начало чтения запроса SaveVariance2");
                using (NpgsqlDataReader reader = cmd.ExecuteReader())
                {
                    int iter = 0;
                    int stabVar = 0;
                    MethodReadProd methodRead = new MethodReadProd(); // Создаем экземпляр класса MethodRead

                    while (reader.Read())
                    {
                        // Заменяем блок инициализации переменных вызовом метода ProcessData
                        methodRead.ProcessData(reader, true, true, true); // Передаем reader и boolOption=true в метод

                        string outputString;

                        if (methodRead.Unspecified)
                        {
                            // Последние три символа productName
                            string lastThreeChars = methodRead.ProductName.Substring(methodRead.ProductName.Length - 3);
                            outputString = $"{methodRead.ProductName} {methodRead.Length}x{methodRead.Width}x{methodRead.Depth}({lastThreeChars})";
                        }
                        else
                        {
                            outputString = $"{methodRead.ProductName} {methodRead.Length}x{methodRead.Width}x{methodRead.Depth}";
                        }
                        if (prodList.Contains(outputString))
                        {
                            if (methodRead.ShiftNum != stabVar)
                            {
                                if (stabVar == 0)
                                {
                                    stabVar = methodRead.ShiftNum;
                                }
                                else
                                {
                                    stabVar = methodRead.ShiftNum;
                                    iter = 0;
                                    firstRow = firstRow + 29;
                                }
                            }
                            iter++;
                            // Теперь передаем methodRead вместо отдельных переменных в метод для сохранения в Excel
                            FillExcel filler = new FillExcel();
                            filler.WriteToExcel2(workbook, methodRead.ShiftDate, methodRead.ShiftNum, methodRead.OnePackVolume, outputString, methodRead.AvgDensity,
                            methodRead.CountGoodProd, methodRead.GoodProdVolume, methodRead.GoodProdWeight, methodRead.CountLowQual, methodRead.LowQualVolume,
                            methodRead.LowQualWeight, methodRead.RejectVolume, methodRead.RejectWeight, firstRow, iter, methodRead.Unspecified,
                            VarDate1, VarDate2);

                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                ProjectLogger.LogException("Ошибка в Чтение запроса. Выгрузка по продуктам за сутки", ex);
            }
        }

        //  -----------------------------------------------------------------------------------------
        // Чтение запроса. Выгрузка по двидам брака
        private void SaveVariance3(XLWorkbook workbook, NpgsqlCommand cmd, List<string> prodList, int firstRow)
        {
            try
            {
                ProjectLogger.LogDebug("Начало чтения запроса SaveVariance3");
                using (NpgsqlDataReader reader = cmd.ExecuteReader())
                {
                    int iter = 0;
                    MethodReadProd methodRead = new MethodReadProd(); // Создаем экземпляр класса MethodRead

                    while (reader.Read())
                    {
                        methodRead.ProcessData(reader, false, false, false); // Передаем reader и параметры в метод

                        iter++;

                        // Теперь передаем значения в метод для сохранения в Excel
                        FillExcel filler = new FillExcel();
                        filler.WriteToExcel3(workbook, methodRead.defectName, methodRead.packCount, methodRead.defectVolume,
                            methodRead.defectWeight, firstRow, iter, VarDate1, VarDate2);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                ProjectLogger.LogException("Ошибка в Чтение запроса. Выгрузка по двидам брака", ex);
            }
        }

        //  -----------------------------------------------------------------------------------------
        // Чтение запроса. Выгрузка по видам брака разбивка по сменам
        private void SaveVariance7(XLWorkbook workbook, NpgsqlCommand cmd, List<string> prodList, int firstRow)
        {
            try
            {
                ProjectLogger.LogDebug("Начало чтения запроса SaveVariance7");
                using (NpgsqlDataReader reader = cmd.ExecuteReader())
                {
                    int iter = 0;
                    int stabVar = 0;
                    MethodReadProd methodRead = new MethodReadProd(); // Создаем экземпляр класса MethodRead

                    while (reader.Read())
                    {
                        methodRead.ProcessData(reader, false, false, true); // Передаем reader и параметры в метод

                        if (methodRead.shiftNum != stabVar)
                        {
                            if (stabVar == 0)
                            {
                                stabVar = methodRead.shiftNum;
                            }
                            else
                            {
                                stabVar = methodRead.shiftNum;
                                iter = 0;
                                firstRow = firstRow + 27;
                            }
                        }

                        iter++;

                        // Теперь передаем значения в метод для сохранения в Excel
                        FillExcel filler = new FillExcel();
                        filler.WriteToExcel3(workbook, methodRead.defectName, methodRead.packCount, methodRead.defectVolume,
                            methodRead.defectWeight, firstRow, iter, VarDate1, VarDate2, stabVar);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                ProjectLogger.LogException("Ошибка в Чтение запроса. Выгрузка по видам брака разбивка по сменам", ex);
            }
        }

        //  -----------------------------------------------------------------------------------------
        // Чтение запроса. Выгрузка по продуктам за предыдущую смену день
        private void SaveVariance4(XLWorkbook workbook, NpgsqlCommand cmd, List<string> prodList, int firstRow)
        {
            try
            {
                ProjectLogger.LogDebug("Начало чтения запроса SaveVariance4");
                using (NpgsqlDataReader reader = cmd.ExecuteReader())
                {
                    int iter = 0;
                    MethodReadProd methodRead = new MethodReadProd(); // Создаем экземпляр класса MethodRead

                    while (reader.Read())
                    {
                        // Заменяем блок инициализации переменных вызовом метода ProcessData
                        methodRead.ProcessData(reader, true, true, true); // Передаем reader и boolOption=true в метод

                        string outputString;

                        if (methodRead.Unspecified)
                        {
                            // Последние три символа productName
                            string lastThreeChars = methodRead.ProductName.Substring(methodRead.ProductName.Length - 3);
                            outputString = $"{methodRead.ProductName} {methodRead.Length}x{methodRead.Width}x{methodRead.Depth}({lastThreeChars})";
                        }
                        else
                        {
                            outputString = $"{methodRead.ProductName} {methodRead.Length}x{methodRead.Width}x{methodRead.Depth}";
                        }
                        if (prodList.Contains(outputString))
                        {
                            iter++;
                            // Теперь передаем methodRead вместо отдельных переменных в метод для сохранения в Excel
                            FillExcel filler = new FillExcel();
                            filler.WriteToExcel5(workbook, methodRead.ShiftDate, methodRead.ShiftNum, methodRead.OnePackVolume, outputString, methodRead.AvgDensity,
                                methodRead.CountGoodProd, methodRead.GoodProdVolume, methodRead.GoodProdWeight, methodRead.CountLowQual, methodRead.LowQualVolume,
                                methodRead.LowQualWeight, methodRead.RejectVolume, methodRead.RejectWeight, firstRow, iter, methodRead.Unspecified,
                                VarDate1, VarDate2);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                ProjectLogger.LogException("Ошибка в Чтение запроса. Выгрузка по продуктам за предыдущую смену день", ex);
            }
        }

        //  -----------------------------------------------------------------------------------------
        // Чтение запроса. Выгрузка по продуктам за период с разбиением по сменам
        private void SaveVariance5(XLWorkbook workbook, NpgsqlCommand cmd, List<string> prodList, int firstRow)
        {
            try
            {
                ProjectLogger.LogDebug("Начало чтения запроса SaveVariance5");
                using (NpgsqlDataReader reader = cmd.ExecuteReader())
                {
                    int iter = 0;
                    int stabVar = 0;
                    MethodReadProd methodRead = new MethodReadProd(); // Создаем экземпляр класса MethodRead

                    while (reader.Read())
                    {
                        // Заменяем блок инициализации переменных вызовом метода ProcessData
                        methodRead.ProcessData(reader, true, false, true); // Передаем reader и boolOption=true в метод

                        string outputString;

                        if (methodRead.Unspecified)
                        {
                            // Последние три символа productName
                            string lastThreeChars = methodRead.ProductName.Substring(methodRead.ProductName.Length - 3);
                            outputString = $"{methodRead.ProductName} {methodRead.Length}x{methodRead.Width}x{methodRead.Depth}({lastThreeChars})";
                        }
                        else
                        {
                            outputString = $"{methodRead.ProductName} {methodRead.Length}x{methodRead.Width}x{methodRead.Depth}";
                        }
                        if (prodList.Contains(outputString))
                        {
                            if (methodRead.ShiftNum != stabVar)
                            {
                                if (stabVar == 0)
                                {
                                    stabVar = methodRead.ShiftNum;
                                }
                                else
                                {
                                    stabVar = methodRead.ShiftNum;
                                    iter = 0;
                                    firstRow = firstRow + 75;
                                }
                            }
                            iter++;

                            // Теперь передаем methodRead вместо отдельных переменных в метод для сохранения в Excel
                            FillExcel filler = new FillExcel();
                            filler.WriteToExcel6(workbook, methodRead.ShiftNum, outputString, methodRead.AvgDensity, methodRead.CountGoodProd, methodRead.GoodProdVolume, methodRead.GoodProdWeight,
                               methodRead.CountLowQual, methodRead.LowQualVolume, methodRead.LowQualWeight, methodRead.RejectVolume, methodRead.RejectWeight,
                               methodRead.LowQualPercent, methodRead.RejectPercent, firstRow, iter, VarDate1, VarDate2);

                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                ProjectLogger.LogException("Ошибка в Чтение запроса. Выгрузка по продуктам за период с разбиением по сменам", ex);
            }
        }
        //  -----------------------------------------------------------------------------------------
        // Чтение запроса. Выгрузка по остановкам
        private void SaveVariance6(XLWorkbook workbook, NpgsqlCommand cmd, int firstRow)
        {
            try
            {
                ProjectLogger.LogDebug("Начало чтения запроса SaveVariance6");
                using (NpgsqlDataReader reader = cmd.ExecuteReader())
                {
                    int iter = 0;
                    MethodReadLine methodReadLine = new MethodReadLine(); // Создаем экземпляр класса MethodRead

                    while (reader.Read())
                    {
                        // Заменяем блок инициализации переменных вызовом метода ProcessData
                        methodReadLine.ProcessDataLine(reader, 0); // Передаем reader и boolOption=true в метод
                        iter++;

                        // Теперь передаем methodRead вместо отдельных переменных в метод для сохранения в Excel
                        FillExcel filler = new FillExcel();
                        filler.WriteToExcel4(workbook, methodReadLine.TypeIndex, methodReadLine.TypeName, methodReadLine.PlaceName,
                           methodReadLine.DurationStopMin, firstRow, iter, VarDate1, VarDate2);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                ProjectLogger.LogException("Ошибка в Чтение запроса. Выгрузка по продуктам за период с разбиением по сменам", ex);
            }
        }
        //  -----------------------------------------------------------------------------------------
        // Чтение запроса. Выгрузка по остановкам по видам
        private void SaveVariance8(XLWorkbook workbook, NpgsqlCommand cmd, int firstRow)
        {
            try
            {
                ProjectLogger.LogDebug("Начало чтения запроса SaveVariance6");
                using (NpgsqlDataReader reader = cmd.ExecuteReader())
                {
                    int iter = 0;
                    MethodReadLine methodReadLine = new MethodReadLine(); // Создаем экземпляр класса MethodRead

                    while (reader.Read())
                    {
                        // Заменяем блок инициализации переменных вызовом метода ProcessData
                        methodReadLine.ProcessDataLine(reader, 1); // Передаем reader и boolOption=true в метод
                        iter++;

                        // Теперь передаем methodRead вместо отдельных переменных в метод для сохранения в Excel
                        FillExcel filler = new FillExcel();
                        filler.WriteToExcel7(workbook, methodReadLine.TypeIndex, methodReadLine.DurationStopMin,
                            firstRow, iter, VarDate1, VarDate2);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                ProjectLogger.LogException("Ошибка в Чтение запроса. Выгрузка по продуктам за период с разбиением по сменам", ex);
            }
        }
        //  -----------------------------------------------------------------------------------------
        // Чтение запроса. Выгрузка по остановкам
        private void SaveVariance9(XLWorkbook workbook, NpgsqlCommand cmd, int firstRow)
        {
            try
            {
                ProjectLogger.LogDebug("Начало чтения запроса SaveVariance6");
                using (NpgsqlDataReader reader = cmd.ExecuteReader())
                {
                    int iter = 0;
                    MethodReadLine methodReadLine = new MethodReadLine(); // Создаем экземпляр класса MethodRead

                    while (reader.Read())
                    {
                        // Заменяем блок инициализации переменных вызовом метода ProcessData
                        methodReadLine.ProcessDataLine(reader, 2); // Передаем reader и boolOption=true в метод
                        iter++;

                        // Теперь передаем methodRead вместо отдельных переменных в метод для сохранения в Excel
                        FillExcel filler = new FillExcel();
                        filler.WriteToExcel8(workbook, methodReadLine.ShiftNum, methodReadLine.TypeIndex, methodReadLine.TypeName, methodReadLine.PlaceName,
                           methodReadLine.DurationStopMin, firstRow, iter, VarDate1, VarDate2);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                ProjectLogger.LogException("Ошибка в Чтение запроса. Выгрузка по продуктам за период с разбиением по сменам", ex);
            }
        }
        //  -----------------------------------------------------------------------------------------
        // Чтение запроса. Выгрузка по остановкам с разбивкой по видам простоя по номерам смен
        private void SaveVariance10(XLWorkbook workbook, NpgsqlCommand cmd, int firstRow)
        {
            try
            {
                ProjectLogger.LogDebug("Начало чтения запроса SaveVariance6");
                using (NpgsqlDataReader reader = cmd.ExecuteReader())
                {
                    int iter = 0;
                    MethodReadLine methodReadLine = new MethodReadLine(); // Создаем экземпляр класса MethodRead

                    while (reader.Read())
                    {
                        // Заменяем блок инициализации переменных вызовом метода ProcessData
                        methodReadLine.ProcessDataLine(reader, 3); // Передаем reader и boolOption=true в метод
                        iter++;

                        // Теперь передаем methodRead вместо отдельных переменных в метод для сохранения в Excel
                        FillExcel filler = new FillExcel();
                        filler.WriteToExcel9(workbook, methodReadLine.ShiftNum, methodReadLine.TypeIndex, methodReadLine.PlaceName,
                           methodReadLine.DurationStopMin, firstRow, iter, VarDate1, VarDate2);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                ProjectLogger.LogException("Ошибка в Чтение запроса. Выгрузка по продуктам за период с разбиением по сменам", ex);
            }
        }
        //  -----------------------------------------------------------------------------------------
        // Чтение запроса. Выгрузка по остановкам
        private void SaveVariance11(XLWorkbook workbook, NpgsqlCommand cmd, int firstRow)
        {
            try
            {
                ProjectLogger.LogDebug("Начало чтения запроса SaveVariance6");
                using (NpgsqlDataReader reader = cmd.ExecuteReader())
                {
                    int iter = 0;
                    MethodReadLine methodReadLine = new MethodReadLine(); // Создаем экземпляр класса MethodRead

                    while (reader.Read())
                    {
                        // Заменяем блок инициализации переменных вызовом метода ProcessData
                        methodReadLine.ProcessDataLine(reader, 4); // Передаем reader и boolOption=true в метод
                        iter++;

                        // Теперь передаем methodRead вместо отдельных переменных в метод для сохранения в Excel
                        FillExcel filler = new FillExcel();
                        filler.WriteToExcel10(workbook, methodReadLine.ShiftDate, methodReadLine.TypeIndex, methodReadLine.TypeName, methodReadLine.PlaceName,
                           methodReadLine.DurationStopMin, methodReadLine.CommentStop, firstRow, iter, VarDate1, VarDate2);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                ProjectLogger.LogException("Ошибка в Чтение запроса. Выгрузка по продуктам за период с разбиением по сменам", ex);
            }
        }
        //  -----------------------------------------------------------------------------------------
        // Чтение запроса. Выгрузка по остановкам
        private void SaveVariance12(XLWorkbook workbook, NpgsqlCommand cmd, int firstRow)
        {
            try
            {
                ProjectLogger.LogDebug("Начало чтения запроса SaveVariance6");
                using (NpgsqlDataReader reader = cmd.ExecuteReader())
                {
                    int iter = 0;
                    MethodReadLine methodReadLine = new MethodReadLine(); // Создаем экземпляр класса MethodRead

                    while (reader.Read())
                    {
                        // Заменяем блок инициализации переменных вызовом метода ProcessData
                        methodReadLine.ProcessDataLine(reader, 5); // Передаем reader и boolOption=true в метод
                        iter++;

                        // Теперь передаем methodRead вместо отдельных переменных в метод для сохранения в Excel
                        FillExcel filler = new FillExcel();
                        filler.WriteToExcel8(workbook, methodReadLine.ShiftNum, methodReadLine.TypeIndex, methodReadLine.TypeName, methodReadLine.PlaceName,
                           methodReadLine.DurationStopMin, firstRow, iter, VarDate1, VarDate2);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                ProjectLogger.LogException("Ошибка в Чтение запроса. Выгрузка по продуктам за период с разбиением по сменам", ex);
            }
        }
        //  -----------------------------------------------------------------------------------------
        // Чтение запроса. Выгрузка по остановкам
        private void SaveVariance13(XLWorkbook workbook, NpgsqlCommand cmd, int firstRow)
        {
            try
            {
                ProjectLogger.LogDebug("Начало чтения запроса SaveVariance6");
                using (NpgsqlDataReader reader = cmd.ExecuteReader())
                {
                    int iter = 0;
                    MethodReadLine methodReadLine = new MethodReadLine(); // Создаем экземпляр класса MethodRead

                    while (reader.Read())
                    {
                        // Заменяем блок инициализации переменных вызовом метода ProcessData
                        methodReadLine.ProcessDataLine(reader, 6); // Передаем reader и boolOption=true в метод
                        iter++;

                        // Теперь передаем methodRead вместо отдельных переменных в метод для сохранения в Excel
                        FillExcel filler = new FillExcel();
                        filler.WriteToExcel11(workbook, methodReadLine.TypeIndex, methodReadLine.PlaceName,
                           methodReadLine.DurationStopMin, firstRow, iter, VarDate1, VarDate2);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                ProjectLogger.LogException("Ошибка в Чтение запроса. Выгрузка по продуктам за период с разбиением по сменам", ex);
            }
        }
    }

}
