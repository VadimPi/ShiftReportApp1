using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ShiftReportApp1
{
    internal class FillExcel
    {
        // Заполнение excel с даты по дату
        public void WriteToExcel1(XLWorkbook workbook, string outputString, float avgDensity,
            int countPack, float volumeProd, float weight, int lowQualCount, float lowQualVol, float lowQualWeight,
            float rejectVol, float rejectWeight, float percentLowQual, float percentReject, int firstRow, int iter,
            DateTime varDate1 = default, DateTime varDate2 = default)
        {
            var worksheet = workbook.Worksheet(1);
            int row = firstRow + iter;
            if (row > 67 && row < 143)
            {
                row = row + 17;
            }
            else if (row > 143 && row < 218)
            {
                row = row + 34;
            }
            else if (row > 218)
            {
                row = row + 51;
            }
            worksheet.Cell(firstRow - 5, 8).Value = varDate1;
            worksheet.Cell(firstRow - 5, 11).Value = varDate2;
            worksheet.Cell(row, 1).Value = iter + 1;
            worksheet.Cell(row, 3).Value = outputString; // Записать outputString в колонку C

            // Записать остальные значения
            worksheet.Cell(row, 4).Value = avgDensity;
            worksheet.Cell(row, 11).Value = countPack;
            worksheet.Cell(row, 12).Value = volumeProd;
            worksheet.Cell(row, 13).Value = weight;
            worksheet.Cell(row, 15).Value = lowQualCount;
            worksheet.Cell(row, 16).Value = lowQualVol;
            worksheet.Cell(row, 18).Value = lowQualWeight;
            worksheet.Cell(row, 19).Value = rejectVol;
            worksheet.Cell(row, 20).Value = rejectWeight;
            worksheet.Cell(row, 22).Value = percentLowQual;
            worksheet.Cell(row, 23).Value = percentReject;

            for (int i = 11; i <= 20 && i != 14 && i != 17; i++)
            {
                var range = worksheet.Range(worksheet.Cell(firstRow + 1, i), worksheet.Cell(firstRow + 59, i));
                var sum = range.CellsUsed().Select(cell => cell.GetDouble()).Sum();
                worksheet.Cell(firstRow + 60, i).Value = sum;
            }
        }
        //  -----------------------------------------------------------------------------------------
        // Заполнение excel за предыдущие сутки
        public void WriteToExcel2(XLWorkbook workbook, DateTime shiftDate, int shiftNum, float onePackVolume,
            string outputString, float avgDensity, int countPack, float volumeProd, float weight, int lowQualCount,
            float lowQualVol, float lowQualWeight, float rejectVol, float rejectWeight, int firstRow, int iter, bool unspecified,
            DateTime varDate1 = default, DateTime varDate2 = default)
        {
            var worksheet = workbook.Worksheet(1);
            int row = firstRow + iter;

            worksheet.Cell(firstRow - 5, 7).Value = shiftDate;
            worksheet.Cell(firstRow - 5, 17).Value = shiftNum;
            worksheet.Cell(row, 1).Value = iter;
            // Значения смены 2,16 25,16
            worksheet.Cell(row, 2).Value = outputString; // Записать outputString в колонку C
            worksheet.Cell(row, 4).Value = onePackVolume;
            worksheet.Cell(row, 5).Value = avgDensity;
            worksheet.Cell(row, 8).Value = countPack;
            worksheet.Cell(row, 10).Value = volumeProd;
            worksheet.Cell(row, 11).Value = weight;
            worksheet.Cell(row, 13).Value = lowQualCount;
            worksheet.Cell(row, 15).Value = lowQualVol;
            worksheet.Cell(row, 17).Value = lowQualWeight;
            worksheet.Cell(row, 19).Value = rejectVol;
            worksheet.Cell(row, 20).Value = rejectWeight;

            if (unspecified) // подсчет пересорта
            {
                worksheet.Cell(firstRow + 15, 22).Value = worksheet.Cell(firstRow + 15, 22).GetDouble() + volumeProd;
                worksheet.Cell(99, 22).Value = worksheet.Cell(99, 22).GetDouble() + volumeProd;
            }

            if (firstRow + iter + 60 > 80) // запись в сутки
            {
                worksheet.Cell(row + 46, 1).Value = iter;
                worksheet.Cell(row + 46, 2).Value = shiftNum;
                worksheet.Cell(row + 46, 3).Value = outputString; // Записать outputString в колонку C
                worksheet.Cell(row + 46, 4).Value = avgDensity;
                worksheet.Cell(row + 46, 10).Value = onePackVolume;
                worksheet.Cell(row + 46, 11).Value = countPack;
                worksheet.Cell(row + 46, 12).Value = volumeProd;
                worksheet.Cell(row + 46, 13).Value = weight;
                worksheet.Cell(row + 46, 15).Value = lowQualCount;
                worksheet.Cell(row + 46, 16).Value = lowQualVol;
                worksheet.Cell(row + 46, 18).Value = lowQualWeight;
                worksheet.Cell(row + 46, 19).Value = rejectVol;
                worksheet.Cell(row + 46, 20).Value = rejectWeight;
            }
            else
            {
                worksheet.Cell(row + 66, 1).Value = iter;
                worksheet.Cell(row + 66, 2).Value = shiftNum;
                worksheet.Cell(row + 66, 3).Value = outputString; // Записать outputString в колонку C
                worksheet.Cell(row + 66, 4).Value = avgDensity;
                worksheet.Cell(row + 66, 10).Value = onePackVolume;
                worksheet.Cell(row + 66, 11).Value = countPack;
                worksheet.Cell(row + 66, 12).Value = volumeProd;
                worksheet.Cell(row + 66, 13).Value = weight;
                worksheet.Cell(row + 66, 15).Value = lowQualCount;
                worksheet.Cell(row + 66, 16).Value = lowQualVol;
                worksheet.Cell(row + 66, 18).Value = lowQualWeight;
                worksheet.Cell(row + 66, 19).Value = rejectVol;
                worksheet.Cell(row + 66, 20).Value = rejectWeight;
            }
            // подсчет за смену
            worksheet.Cell(firstRow + 15, 8).Value = worksheet.Cell(firstRow + 15, 8).GetDouble() + countPack;
            worksheet.Cell(firstRow + 15, 10).Value = worksheet.Cell(firstRow + 15, 10).GetDouble() + volumeProd;
            worksheet.Cell(firstRow + 15, 11).Value = worksheet.Cell(firstRow + 15, 11).GetDouble() + weight;
            worksheet.Cell(firstRow + 15, 13).Value = worksheet.Cell(firstRow + 15, 13).GetDouble() + lowQualCount;
            worksheet.Cell(firstRow + 15, 15).Value = worksheet.Cell(firstRow + 15, 15).GetDouble() + lowQualVol;
            worksheet.Cell(firstRow + 15, 17).Value = worksheet.Cell(firstRow + 15, 17).GetDouble() + lowQualWeight;
            worksheet.Cell(firstRow + 15, 19).Value = worksheet.Cell(firstRow + 15, 19).GetDouble() + rejectVol;
            worksheet.Cell(firstRow + 15, 20).Value = worksheet.Cell(firstRow + 15, 20).GetDouble() + rejectWeight;

            worksheet.Cell(firstRow + 16, 4).Value = worksheet.Cell(firstRow + 16, 4).GetDouble() + weight + lowQualWeight + rejectWeight;
            worksheet.Cell(firstRow + 16, 12).Value = worksheet.Cell(firstRow + 15, 17).GetDouble() / worksheet.Cell(firstRow + 16, 4).GetDouble();
            worksheet.Cell(firstRow + 16, 17).Value = worksheet.Cell(firstRow + 15, 20).GetDouble() / worksheet.Cell(firstRow + 16, 4).GetDouble();
            worksheet.Cell(firstRow + 16, 20).Value = worksheet.Cell(firstRow + 15, 22).GetDouble() / worksheet.Cell(firstRow + 16, 4).GetDouble();

            // подсчет за сутки
            worksheet.Cell(97, 11).Value = worksheet.Cell(21, 8).GetDouble();
            worksheet.Cell(97, 12).Value = worksheet.Cell(21, 10).GetDouble();
            worksheet.Cell(97, 13).Value = worksheet.Cell(21, 11).GetDouble();
            worksheet.Cell(97, 15).Value = worksheet.Cell(21, 13).GetDouble();
            worksheet.Cell(97, 16).Value = worksheet.Cell(21, 15).GetDouble();
            worksheet.Cell(97, 18).Value = worksheet.Cell(21, 17).GetDouble();
            worksheet.Cell(97, 19).Value = worksheet.Cell(21, 19).GetDouble();
            worksheet.Cell(97, 20).Value = worksheet.Cell(21, 20).GetDouble();
            worksheet.Cell(98, 11).Value = worksheet.Cell(51, 8).GetDouble();
            worksheet.Cell(98, 12).Value = worksheet.Cell(51, 10).GetDouble();
            worksheet.Cell(98, 13).Value = worksheet.Cell(51, 11).GetDouble();
            worksheet.Cell(98, 15).Value = worksheet.Cell(51, 13).GetDouble();
            worksheet.Cell(98, 16).Value = worksheet.Cell(51, 15).GetDouble();
            worksheet.Cell(98, 18).Value = worksheet.Cell(51, 17).GetDouble();
            worksheet.Cell(98, 19).Value = worksheet.Cell(51, 19).GetDouble();
            worksheet.Cell(98, 20).Value = worksheet.Cell(51, 20).GetDouble();

            worksheet.Cell(99, 11).Value = worksheet.Cell(99, 11).GetDouble() + countPack;
            worksheet.Cell(99, 12).Value = worksheet.Cell(99, 12).GetDouble() + volumeProd;
            worksheet.Cell(99, 13).Value = worksheet.Cell(99, 13).GetDouble() + weight;
            worksheet.Cell(99, 15).Value = worksheet.Cell(99, 15).GetDouble() + lowQualCount;
            worksheet.Cell(99, 16).Value = worksheet.Cell(99, 16).GetDouble() + lowQualVol;
            worksheet.Cell(99, 18).Value = worksheet.Cell(99, 18).GetDouble() + lowQualWeight;
            worksheet.Cell(99, 19).Value = worksheet.Cell(99, 19).GetDouble() + rejectVol;
            worksheet.Cell(99, 20).Value = worksheet.Cell(99, 20).GetDouble() + rejectWeight;

            worksheet.Cell(100, 4).Value = worksheet.Cell(100, 4).GetDouble() + weight + lowQualWeight + rejectWeight;
            worksheet.Cell(100, 8).Value = worksheet.Cell(100, 8).GetDouble() + volumeProd + lowQualVol + rejectVol;
            worksheet.Cell(100, 12).Value = worksheet.Cell(99, 18).GetDouble() / worksheet.Cell(100, 4).GetDouble();
            worksheet.Cell(100, 17).Value = worksheet.Cell(99, 20).GetDouble() / worksheet.Cell(100, 4).GetDouble();
            worksheet.Cell(100, 20).Value = worksheet.Cell(99, 22).GetDouble() / worksheet.Cell(100, 4).GetDouble();
        }
        //  -----------------------------------------------------------------------------------------
        // Разбивка по видам брака по сменам за период
        public void WriteToExcel3(XLWorkbook workbook, string defectName, int packCount, float defectVolume, float defectWeight,
            int firstRow, int iter, DateTime varDate1 = default, DateTime varDate2 = default, int shiftNum = 0)
        {
            var worksheet = workbook.Worksheet(1);
            int row = firstRow + iter;

            if (shiftNum != 0)
            {
                worksheet.Cell(firstRow - 4, 4).Value = shiftNum; // Записать смену если разбивка по сменам
            }

            worksheet.Cell(firstRow - 3, 8).Value = varDate1;
            worksheet.Cell(firstRow - 3, 11).Value = varDate2;
            worksheet.Cell(row, 1).Value = iter;
            worksheet.Cell(row, 2).Value = defectName;
            worksheet.Cell(row, 11).Value = packCount;
            worksheet.Cell(row, 14).Value = defectVolume;
            worksheet.Cell(row, 18).Value = defectWeight;
            worksheet.Cell(21, 11).Value = worksheet.Cell(21, 11).GetDouble() + packCount;
            worksheet.Cell(21, 14).Value = worksheet.Cell(21, 14).GetDouble() + defectVolume;
            worksheet.Cell(21, 18).Value = worksheet.Cell(21, 18).GetDouble() + defectWeight;
            worksheet.Cell(row, 9).Value = defectWeight / worksheet.Cell(21, 18).GetDouble() * 100;
        }
        //  -----------------------------------------------------------------------------------------
        // За предыдущую смену
        public void WriteToExcel5(XLWorkbook workbook, DateTime shiftDate, int shiftNum, float onePackVolume, string outputString,
            float avgDensity, int countPack, float volumeProd, float weight, int lowQualCount, float lowQualVol, float lowQualWeight,
            float rejectVol, float rejectWeight, int firstRow, int iter, bool unspecified, DateTime varDate1 = default,
            DateTime varDate2 = default)
        {
            var worksheet = workbook.Worksheet(1);
            int row = firstRow + iter;

            worksheet.Cell(firstRow - 5, 7).Value = shiftDate;
            worksheet.Cell(firstRow - 5, 17).Value = shiftNum;
            worksheet.Cell(row, 1).Value = iter;
            // Значения смены 2,16 25,16
            worksheet.Cell(row, 2).Value = outputString; // Записать outputString в колонку C
            worksheet.Cell(row, 4).Value = onePackVolume;
            worksheet.Cell(row, 5).Value = avgDensity;
            worksheet.Cell(row, 8).Value = countPack;
            worksheet.Cell(row, 10).Value = volumeProd;
            worksheet.Cell(row, 11).Value = weight;
            worksheet.Cell(row, 13).Value = lowQualCount;
            worksheet.Cell(row, 15).Value = lowQualVol;
            worksheet.Cell(row, 17).Value = lowQualWeight;
            worksheet.Cell(row, 19).Value = rejectVol;
            worksheet.Cell(row, 20).Value = rejectWeight;
            if (unspecified == true)
            {
                worksheet.Cell(firstRow + 15, 22).Value = worksheet.Cell(firstRow + 15, 22).GetDouble() + volumeProd;
            }
            worksheet.Cell(firstRow + 15, 8).Value = worksheet.Cell(firstRow + 15, 8).GetDouble() + countPack;
            worksheet.Cell(firstRow + 15, 10).Value = worksheet.Cell(firstRow + 15, 10).GetDouble() + volumeProd;
            worksheet.Cell(firstRow + 15, 11).Value = worksheet.Cell(firstRow + 15, 11).GetDouble() + weight;
            worksheet.Cell(firstRow + 15, 13).Value = worksheet.Cell(firstRow + 15, 13).GetDouble() + lowQualCount;
            worksheet.Cell(firstRow + 15, 15).Value = worksheet.Cell(firstRow + 15, 15).GetDouble() + lowQualVol;
            worksheet.Cell(firstRow + 15, 17).Value = worksheet.Cell(firstRow + 15, 17).GetDouble() + lowQualWeight;
            worksheet.Cell(firstRow + 15, 19).Value = worksheet.Cell(firstRow + 15, 19).GetDouble() + rejectVol;
            worksheet.Cell(firstRow + 15, 20).Value = worksheet.Cell(firstRow + 15, 20).GetDouble() + rejectWeight;

            worksheet.Cell(firstRow + 16, 4).Value = worksheet.Cell(22, 4).GetDouble() + weight + lowQualWeight + rejectWeight;
            worksheet.Cell(firstRow + 16, 12).Value = worksheet.Cell(firstRow + 15, 17).GetDouble() / worksheet.Cell(firstRow + 16, 4).GetDouble();
            worksheet.Cell(firstRow + 16, 17).Value = worksheet.Cell(firstRow + 15, 20).GetDouble() / worksheet.Cell(firstRow + 16, 4).GetDouble();
            worksheet.Cell(firstRow + 16, 20).Value = worksheet.Cell(firstRow + 15, 22).GetDouble() / worksheet.Cell(firstRow + 16, 4).GetDouble();
        }
        //  -----------------------------------------------------------------------------------------
        // Заполнение excel с даты по дату с разбивкой по сменам
        public void WriteToExcel6(XLWorkbook workbook, int shiftNum, string outputString, float avgDensity,
            int countPack, float volumeProd, float weight, int lowQualCount, float lowQualVol, float lowQualWeight,
            float rejectVol, float rejectWeight, float percentLowQual, float percentReject, int firstRow, int iter,
            DateTime varDate1 = default, DateTime varDate2 = default)
        {
            var worksheet = workbook.Worksheet(1);
            int row = firstRow + iter;
            worksheet.Cell(firstRow - 5, 8).Value = varDate1;
            worksheet.Cell(firstRow - 5, 11).Value = varDate2;
            worksheet.Cell(row, 1).Value = iter;
            worksheet.Cell(row, 2).Value = shiftNum;
            worksheet.Cell(row, 3).Value = outputString; // Записать outputString в колонку C
            worksheet.Cell(row, 4).Value = avgDensity;
            worksheet.Cell(row, 11).Value = countPack;
            worksheet.Cell(row, 12).Value = volumeProd;
            worksheet.Cell(row, 13).Value = weight;
            worksheet.Cell(row, 15).Value = lowQualCount;
            worksheet.Cell(row, 16).Value = lowQualVol;
            worksheet.Cell(row, 18).Value = lowQualWeight;
            worksheet.Cell(row, 19).Value = rejectVol;
            worksheet.Cell(row, 20).Value = rejectWeight;
            worksheet.Cell(row, 22).Value = percentLowQual;
            worksheet.Cell(row, 23).Value = percentReject;

            for (int i = 11; i <= 20 && i != 14 && i != 17; i++)
            {
                var range = worksheet.Range(worksheet.Cell(firstRow + 1, i), worksheet.Cell(firstRow + 59, i));
                var sum = range.CellsUsed().Select(cell => cell.GetDouble()).Sum();
                worksheet.Cell(firstRow + 60, i).Value = sum;
            }
        }
        //  -----------------------------------------------------------------------------------------
        // Заполнение excel с даты по дату по остановкам
        public void WriteToExcel4(XLWorkbook workbook, string TypeIndex, string TypeName, string PlaceName, int DurationStopMin, int firstRow, int iter,
            DateTime varDate1 = default, DateTime varDate2 = default)
        {
            var worksheet = workbook.Worksheet(1);
            int row = firstRow + iter;
            worksheet.Cell(firstRow - 4, 7).Value = varDate1;
            worksheet.Cell(firstRow - 4, 9).Value = varDate2;
            worksheet.Cell(row, 1).Value = iter;
            worksheet.Cell(row, 5).Value = TypeIndex;
            worksheet.Cell(row, 7).Value = TypeName;
            worksheet.Cell(row, 9).Value = PlaceName;
            worksheet.Cell(row, 13).Value = DurationStopMin;

            worksheet.Cell(68, 13).Value = worksheet.Cell(68, 13).GetDouble() + DurationStopMin;
        }
        //  -----------------------------------------------------------------------------------------
        // Заполнение excel с даты по дату по видам остановок
        public void WriteToExcel7(XLWorkbook workbook, string TypeIndex, int DurationStopMin, int firstRow, int iter,
            DateTime varDate1 = default, DateTime varDate2 = default)
        {
            var worksheet = workbook.Worksheet(1);
            int row = firstRow + iter;
            worksheet.Cell(firstRow - 4, 7).Value = varDate1;
            worksheet.Cell(firstRow - 4, 9).Value = varDate2;
            worksheet.Cell(row, 1).Value = iter;
            worksheet.Cell(row, 5).Value = TypeIndex;
            worksheet.Cell(row, 13).Value = DurationStopMin;

            worksheet.Cell(68, 13).Value = worksheet.Cell(68, 13).GetDouble() + DurationStopMin;
        }
        //  -----------------------------------------------------------------------------------------
        // Заполнение excel с даты по дату по простоям с разбивкой по номерам смен
        public void WriteToExcel8(XLWorkbook workbook, int ShiftNum, string TypeIndex, string TypeName, string PlaceName, int DurationStopMin, int firstRow, int iter,
            DateTime varDate1 = default, DateTime varDate2 = default)
        {
            var worksheet = workbook.Worksheet(1);
            int row = firstRow + iter;
            worksheet.Cell(firstRow - 4, 7).Value = varDate1;
            worksheet.Cell(firstRow - 4, 9).Value = varDate2;
            worksheet.Cell(row, 1).Value = iter;
            worksheet.Cell(row, 4).Value = ShiftNum;
            worksheet.Cell(row, 5).Value = TypeIndex;
            worksheet.Cell(row, 7).Value = TypeName;
            worksheet.Cell(row, 9).Value = PlaceName;
            worksheet.Cell(row, 13).Value = DurationStopMin;

            worksheet.Cell(68, 13).Value = worksheet.Cell(68, 13).GetDouble() + DurationStopMin;
        }
        //  -----------------------------------------------------------------------------------------
        // Заполнение excel с даты по дату по простоям с разбивкой по номерам смен и по типам остановок
        public void WriteToExcel9(XLWorkbook workbook, int ShiftNum, string TypeIndex, string PlaceName, int DurationStopMin, int firstRow, int iter,
            DateTime varDate1 = default, DateTime varDate2 = default)
        {
            var worksheet = workbook.Worksheet(1);
            int row = firstRow + iter;
            worksheet.Cell(firstRow - 4, 7).Value = varDate1;
            worksheet.Cell(firstRow - 4, 9).Value = varDate2;
            worksheet.Cell(row, 1).Value = iter;
            worksheet.Cell(row, 4).Value = ShiftNum;
            worksheet.Cell(row, 5).Value = TypeIndex;
            worksheet.Cell(row, 9).Value = PlaceName;
            worksheet.Cell(row, 13).Value = DurationStopMin;

            worksheet.Cell(68, 13).Value = worksheet.Cell(68, 13).GetDouble() + DurationStopMin;
        }
        //  -----------------------------------------------------------------------------------------
        // Заполнение excel с даты по дату по простоям с комментариями
        public void WriteToExcel10(XLWorkbook workbook, DateTime ShiftDate, string TypeIndex, string TypeName, string PlaceName, int DurationStopMin,
            string CommentStop, int firstRow, int iter, DateTime varDate1 = default, DateTime varDate2 = default)
        {
            var worksheet = workbook.Worksheet(1);
            int row = firstRow + iter;
            worksheet.Cell(firstRow - 4, 7).Value = varDate1;
            worksheet.Cell(firstRow - 4, 9).Value = varDate2;
            worksheet.Cell(row, 1).Value = iter;
            worksheet.Cell(row, 2).Value = ShiftDate;
            worksheet.Cell(row, 5).Value = TypeIndex;
            worksheet.Cell(row, 7).Value = TypeName;
            worksheet.Cell(row, 9).Value = PlaceName;
            worksheet.Cell(row, 13).Value = DurationStopMin;
            worksheet.Cell(row, 14).Value = CommentStop;

            worksheet.Cell(68, 13).Value = worksheet.Cell(68, 13).GetDouble() + DurationStopMin;
        }
        //  -----------------------------------------------------------------------------------------
        // Заполнение excel с даты по дату по простоям по месту
        public void WriteToExcel11(XLWorkbook workbook, string TypeIndex, string PlaceName, int DurationStopMin,
            int firstRow, int iter, DateTime varDate1 = default, DateTime varDate2 = default)
        {
            var worksheet = workbook.Worksheet(1);
            int row = firstRow + iter;
            worksheet.Cell(firstRow - 4, 7).Value = varDate1;
            worksheet.Cell(firstRow - 4, 9).Value = varDate2;
            worksheet.Cell(row, 1).Value = iter;
            worksheet.Cell(row, 5).Value = TypeIndex;
            worksheet.Cell(row, 9).Value = PlaceName;
            worksheet.Cell(row, 13).Value = DurationStopMin;

            worksheet.Cell(68, 13).Value = worksheet.Cell(68, 13).GetDouble() + DurationStopMin;
        }
        //  -----------------------------------------------------------------------------------------
    }
}
