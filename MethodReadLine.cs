using DocumentFormat.OpenXml.Bibliography;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ShiftReportApp1
{
    internal class MethodReadLine
    {
        public string PlaceName { get; set; }
        public string Section { get; set; }
        public int ShiftDay { get; set; }
        public int ShiftNum { get; set; }
        public DateTime ShiftDate { get; set; }
        public string TypeName { get; set; }
        public string TypeIndex { get; set; }
        public DateTime StopFirstTime { get; set; }
        public DateTime StopEndTime { get; set; }
        public string CommentStop { get; set; }
        public int DurationStopMin { get; set; }
        public DateTime RecordStopReport { get; set; }

        // ----------------------------------------------------------------------------------------------------------------------------
        public void ProcessDataLine(IDataReader reader, int ReadOption)
        {
            if (ReadOption == 0)
            {
                // Блок 1
                TypeIndex = $"{reader["Типы_простоев"]}";
                TypeName = $"{reader["Название_остановки"]}";
                PlaceName = $"{reader["Место"]}";
                DurationStopMin = (int)reader["Время_простоя"];
            }
            if (ReadOption == 1)
            {
                // Блок 2
                TypeIndex = $"{reader["Типы_простоев"]}";
                DurationStopMin = (int)reader["Время_простоя"];
            }
            if (ReadOption == 2)
            {
                // Блок 3
                ShiftNum = (int)reader["Номер_смены"];
                TypeIndex = $"{reader["Типы_простоев"]}";
                TypeName = $"{reader["Название_остановки"]}";
                PlaceName = $"{reader["Место"]}";
                DurationStopMin = (int)reader["Время_простоя"];
            }
            if (ReadOption == 3)
            {
                // Блок 4
                ShiftNum = (int)reader["Номер_смены"];
                TypeIndex = $"{reader["Типы_простоев"]}";
                PlaceName = $"{reader["Место"]}";
                DurationStopMin = (int)reader["Время_простоя"];
            }
            if (ReadOption == 4)
            {
                // Блок 5
                ShiftDate = (DateTime)reader["Дата"];
                TypeIndex = $"{reader["Типы_простоев"]}";
                TypeName = $"{reader["Название_остановки"]}";
                PlaceName = $"{reader["Место"]}";
                DurationStopMin = (int)reader["Время_простоя"];
                CommentStop = $"{reader["Комментарий"]}";
            }
            if (ReadOption == 5)
            {
                // Блок 6
                ShiftNum = (int)reader["Номер_смены"];
                TypeIndex = $"{reader["Типы_простоев"]}";
                TypeName = $"{reader["Название_остановки"]}";
                PlaceName = $"{reader["Место"]}";
                DurationStopMin = (int)reader["Время_простоя"];
            }
            if (ReadOption == 6)
            {
                // Блок 7
                TypeIndex = $"{reader["Типы_простоев"]}";
                PlaceName = $"{reader["Место"]}";
                DurationStopMin = (int)reader["Время_простоя"];
            }
        }
    }
}

