using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ShiftReportApp1
{
    internal class MethodReadProd
    {
        public string ProductName { get; set; }
        public string Depth { get; set; }
        public string Length { get; set; }
        public string Width { get; set; }
        public bool Unspecified { get; set; }
        public float AvgDensity { get; set; }
        public int CountGoodProd { get; set; }
        public float GoodProdVolume { get; set; }
        public float GoodProdWeight { get; set; }
        public int CountLowQual { get; set; }
        public float LowQualVolume { get; set; }
        public float LowQualWeight { get; set; }
        public float RejectVolume { get; set; }
        public float RejectWeight { get; set; }
        public float LowQualPercent { get; set; }
        public float RejectPercent { get; set; }
        public DateTime ShiftDate { get; set; }
        public int ShiftNum { get; set; }
        public float OnePackVolume { get; set; }
        public string OutputString { get; set; }

        public int shiftNum { get; set; }
        public string defectName { get; set; }
        public int packCount { get; set; }
        public float defectVolume { get; set; }
        public float defectWeight { get; set; }


        // ----------------------------------------------------------------------------------------------------------------------------
        public void ProcessData(IDataReader reader, bool prodDefectOption, bool dayOption, bool shiftOption)
        {
            if (prodDefectOption)
            {
                // Блок 1
                ProductName = $"{reader["product_name"]}";
                Depth = $"{reader["prod_depth"]}";
                Length = $"{reader["length"]}";
                Width = $"{reader["width"]}";
                Unspecified = (bool)reader["unspecified"];
                AvgDensity = (float)reader["avg_density_avg"];
                CountGoodProd = (int)reader["pack_count_sum"];
                GoodProdVolume = (float)reader["volume_prod_sum"];
                GoodProdWeight = (float)reader["weight_sum"];
                CountLowQual = (int)reader["low_qual_count"];
                LowQualVolume = (float)reader["low_qual_vol"];
                LowQualWeight = (float)reader["low_qual_weight"];
                RejectVolume = (float)reader["reject_vol"];
                RejectWeight = (float)reader["reject_weight"];

                // Блок 2
                if (dayOption)
                {
                    ShiftDate = (DateTime)reader["shift_date"];
                    OnePackVolume = (float)reader["volume_pack,"];
                }
                else
                {
                    // Блок 3
                    LowQualPercent = (float)reader["percent_low_qual"];
                    RejectPercent = (float)reader["percent_reject"];
                }
                if (shiftOption)
                {
                    ShiftNum = (int)reader["shift_num"];
                }
            }
            else
            {
                defectName = $"{reader["Дефекты"]}";
                packCount = (int)reader["Количество_упаковок"];
                defectVolume = (float)reader["Объем"];
                defectWeight = (float)reader["Вес"];

                if (shiftOption)
                {
                    shiftNum = (int)reader["Смена"];
                }
            }
        }
    }
}
