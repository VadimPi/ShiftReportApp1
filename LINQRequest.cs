using DocumentFormat.OpenXml.Bibliography;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Windows.Forms;

namespace ShiftReportApp1
{
    public class ItemReportList
    {
        public string ProductNames { get; set; }
        public bool Unspecifies { get; set; }
        public int Depth { get; set; }
        public int Length { get; set; }
        public int Width { get; set; }
    }
    public class LINQRequest
    {
        public DataTable ExtractProduct(int getMethod, DateTime varDate1, DateTime varDate2, List<int> shiftDays, List<int> shiftNumbs, List<string> stopCategoryes)
        {
            try
            {
                ProjectLogger.LogDebug("Начало ExtractProduct (LINQRequest)");
                if (getMethod == 0) // отчет по качеству за период с даты по дату
                {
                    using (var dbContext = new ShiftReportDbContext())
                    {
                        var prodQualityReports =
                        from pqr in dbContext.ProdQualityReport
                        join pc in dbContext.ProductCategories on pqr.Product equals pc.ProductID
                        join sr in dbContext.ShiftReport on pqr.Report equals sr.ShiftReportID
                        join pdr in dbContext.ProdDefectReport on pqr.PQReportID equals pdr.ProductReport into defectReports
                        from pdr in defectReports.DefaultIfEmpty() // Выполняем LEFT JOIN
                        where sr.ShiftDate >= varDate1 && sr.ShiftDate <= varDate2
                        group new { pqr, pc, pdr } by new
                        {
                            pc.ProductName,
                            pqr.Unspecified,
                            pqr.ProdDepth,
                            pqr.ProdLength,
                            pqr.ProdWidth
                        } into grp
                        orderby grp.Sum(x => x.pqr.Weight) descending
                        select new
                        {
                            ProductName = grp.Key.ProductName,
                            Unspecified = grp.Key.Unspecified,
                            ProdDepth = grp.Key.ProdDepth,
                            Length = grp.Key.ProdLength,
                            Width = grp.Key.ProdWidth,
                            AvgVolumePack = grp.Average(x => x.pqr.VolumePack),
                            AvgDensityAvg = grp.Average(x => x.pqr.AvgDensity),
                            PackCountSum = grp.Sum(x => x.pqr.PackCount),
                            VolumeProdSum = grp.Sum(x => x.pqr.VolumeProduct),
                            WeightSum = grp.Sum(x => x.pqr.Weight),
                            AvgDensityOS = grp.Average(x => x.pdr.DefectDensity),
                            LowQualCount = grp.Sum(x => x.pdr.DefectType >= 1 && x.pdr.DefectType <= 7 ? x.pdr.DefectPackCount : 0),
                            LowQualVol = grp.Sum(x => x.pdr.DefectType >= 1 && x.pdr.DefectType <= 7 ? x.pdr.DefectVolume : 0),
                            LowQualWeight = grp.Sum(x => x.pdr.DefectType >= 1 && x.pdr.DefectType <= 7 ? x.pdr.DefectWeight : 0),
                            RejectVol = grp.Sum(x => x.pdr.DefectType == 8 ? x.pdr.DefectVolume : 0),
                            RejectWeight = grp.Sum(x => x.pdr.DefectType == 8 ? x.pdr.DefectWeight : 0),
                            PercentLowQual = (grp.Sum(x => x.pdr.DefectType >= 1 && x.pdr.DefectType <= 7 ? x.pdr.DefectWeight : 0) / grp.Sum(x => x.pqr.Weight)) * 100.0,
                            PercentReject = (grp.Sum(x => x.pdr.DefectType == 8 ? x.pdr.DefectWeight : 0) / grp.Sum(x => x.pqr.Weight)) * 100.0
                        };
                        prodQualityReports.ToList();

                        DataTable dataTable = new DataTable();
                        dataTable.Columns.Add("Марка", typeof(string));
                        dataTable.Columns.Add("Неуказанная плт-ть", typeof(bool));
                        dataTable.Columns.Add("Толщина", typeof(int));
                        dataTable.Columns.Add("Длинна", typeof(int));
                        dataTable.Columns.Add("Ширина", typeof(int));
                        dataTable.Columns.Add("Объем пачки", typeof(float));
                        dataTable.Columns.Add("Ср. плотность", typeof(float));
                        dataTable.Columns.Add("Кол-во пачек", typeof(int));
                        dataTable.Columns.Add("Объем", typeof(float));
                        dataTable.Columns.Add("Вес", typeof(float));
                        dataTable.Columns.Add("Средняя плотность ОС", typeof(float));
                        dataTable.Columns.Add("Кол-во пачек ОС", typeof(int));
                        dataTable.Columns.Add("Объем ОС", typeof(float));
                        dataTable.Columns.Add("Вес ОС", typeof(float));
                        dataTable.Columns.Add("Объем обрезь", typeof(float));
                        dataTable.Columns.Add("Вес обрезь", typeof(float));
                        dataTable.Columns.Add("Процент ОС", typeof(double));
                        dataTable.Columns.Add("Процент обрезь", typeof(double));

                        // Заполняем DataTable
                        foreach (var report in prodQualityReports)
                        {
                            DataRow row = dataTable.NewRow();
                            row["Марка"] = report.ProductName;
                            row["Неуказанная плт-ть"] = report.Unspecified;
                            row["Толщина"] = report.ProdDepth;
                            row["Длинна"] = report.Length;
                            row["Ширина"] = report.Width;
                            row["Объем пачки"] = report.AvgVolumePack;
                            row["Ср. плотность"] = report.AvgDensityAvg;
                            row["Кол-во пачек"] = report.PackCountSum;
                            row["Объем"] = report.VolumeProdSum;
                            row["Вес"] = report.WeightSum;
                            row["Средняя плотность ОС"] = report.AvgDensityOS;
                            row["Кол-во пачек ОС"] = report.LowQualCount;
                            row["Объем ОС"] = report.LowQualVol;
                            row["Вес ОС"] = report.LowQualWeight;
                            row["Объем обрезь"] = report.RejectVol;
                            row["Вес обрезь"] = report.RejectWeight;
                            row["Процент ОС"] = report.PercentLowQual;
                            row["Процент обрезь"] = report.PercentReject;
                            // Заполняйте остальные поля аналогично
                            dataTable.Rows.Add(row);
                        }
                        dbContext.Dispose();
                        return dataTable;
                    }
                        
                }
                ProjectLogger.LogDebug("Конец ExtractProduct (LINQRequest)");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                ProjectLogger.LogException("Ошибка в ExtractProduct (LINQRequest)", ex);
            }
            //---------------------------------------------------------------------------------------------------------------------------------------
            if (getMethod == 2) // отчет за предыдущие сутки с разбивкой по сменам
            {
                using (var dbContext = new ShiftReportDbContext())
                {
                    var result = from pqr in dbContext.ProdQualityReport
                                 join pc in dbContext.ProductCategories on pqr.Product equals pc.ProductID
                                 join sr in dbContext.ShiftReport on pqr.Report equals sr.ShiftReportID
                                 join pdr in dbContext.ProdDefectReport on pqr.PQReportID equals pdr.ProductReport into defectReports
                                 from pdr in defectReports.DefaultIfEmpty() // Выполняем LEFT JOIN
                                 where sr.ShiftDate >= varDate1.Date && sr.ShiftDate <= varDate2.Date
                                 group new { pqr, pc, pdr } by new
                                 {
                                     sr.ShiftReportID,
                                     pqr.PQReportID,
                                     sr.ShiftDate,
                                     sr.ShiftNum,
                                     pqr.VolumePack,
                                     pc.ProductName,
                                     pqr.Unspecified,
                                     pqr.Regarding,
                                     pqr.ProdDepth,
                                     pqr.ProdLength,
                                     pqr.ProdWidth
                                 } into grp
                                 orderby grp.Key.PQReportID
                                 select new
                                 {
                                     ShiftDate = grp.Key.ShiftDate,
                                     ShiftNum = grp.Key.ShiftNum,
                                     ProductName = grp.Key.ProductName,
                                     Unspecified = grp.Key.Unspecified,
                                     ProdDepth = grp.Key.ProdDepth,
                                     Length = grp.Key.ProdLength,
                                     Width = grp.Key.ProdWidth,
                                     AvgDensityAvg = grp.Average(x => x.pqr.AvgDensity),
                                     VolumePack = grp.Key.VolumePack,
                                     PackCountSum = grp.Average(x => x.pqr.PackCount),
                                     VolumeProdSum = grp.Average(x => x.pqr.VolumeProduct),
                                     WeightSum = grp.Average(x => x.pqr.Weight),
                                     LowQualCount = grp.Sum(x => x.pdr.DefectType >= 1 && x.pdr.DefectType <= 7 ? x.pdr.DefectPackCount : 0),
                                     LowQualVol = grp.Sum(x => x.pdr.DefectType >= 1 && x.pdr.DefectType <= 7 ? x.pdr.DefectVolume : 0),
                                     LowQualWeight = grp.Sum(x => x.pdr.DefectType >= 1 && x.pdr.DefectType <= 7 ? x.pdr.DefectWeight : 0),
                                     RejectVol = grp.Sum(x => x.pdr.DefectType == 8 ? x.pdr.DefectVolume : 0),
                                     RejectWeight = grp.Sum(x => x.pdr.DefectType == 8 ? x.pdr.DefectWeight : 0),
                                     Regarding = grp.Key.Regarding,
                                     AvgLowQualDensity = grp.Average(x => x.pdr.DefectDensity)
                                 };
                    result.ToList();

                    DataTable dataTable = new DataTable();
                    dataTable.Columns.Add("Дата", typeof(DateTime));
                    dataTable.Columns.Add("Смена", typeof(int));
                    dataTable.Columns.Add("Марка", typeof(string));
                    dataTable.Columns.Add("Неуказанная плт-ть", typeof(bool));
                    dataTable.Columns.Add("Толщина", typeof(int));
                    dataTable.Columns.Add("Длинна", typeof(int));
                    dataTable.Columns.Add("Ширина", typeof(int));
                    dataTable.Columns.Add("Ср. плотность", typeof(float));
                    dataTable.Columns.Add("Объем пачки", typeof(float));
                    dataTable.Columns.Add("Кол-во пачек", typeof(int));
                    dataTable.Columns.Add("Объем", typeof(float));
                    dataTable.Columns.Add("Вес", typeof(float));
                    dataTable.Columns.Add("Кол-во пачек ОС", typeof(int));
                    dataTable.Columns.Add("Объем ОС", typeof(float));
                    dataTable.Columns.Add("Вес ОС", typeof(float));
                    dataTable.Columns.Add("Объем обрезь", typeof(float));
                    dataTable.Columns.Add("Вес обрезь", typeof(float));
                    dataTable.Columns.Add("Пересорт", typeof(bool));
                    dataTable.Columns.Add("Плотность ОС", typeof(float));

                    // Заполняем DataTable
                    foreach (var report in result)
                    {
                        DataRow row = dataTable.NewRow();
                        row["Дата"] = report.ShiftDate;
                        row["Смена"] = report.ShiftNum;
                        row["Марка"] = report.ProductName;
                        row["Неуказанная плт-ть"] = report.Unspecified;
                        row["Толщина"] = report.ProdDepth;
                        row["Длинна"] = report.Length;
                        row["Ширина"] = report.Width;
                        row["Ср. плотность"] = report.AvgDensityAvg;
                        row["Объем пачки"] = report.VolumePack;
                        row["Кол-во пачек"] = report.PackCountSum;
                        row["Объем"] = report.VolumeProdSum;
                        row["Вес"] = report.WeightSum;
                        row["Кол-во пачек ОС"] = report.LowQualCount;
                        row["Объем ОС"] = report.LowQualVol;
                        row["Вес ОС"] = report.LowQualWeight;
                        row["Объем обрезь"] = report.RejectVol;
                        row["Вес обрезь"] = report.RejectWeight;
                        row["Пересорт"] = report.Regarding;
                        row["Плотность ОС"] = report.AvgLowQualDensity;
                        // Заполняйте остальные поля аналогично
                        dataTable.Rows.Add(row);

                    }
                    dbContext.Dispose();
                    return dataTable;
                }
            }
            //---------------------------------------------------------------------------------------------------------------------------------------
            if (getMethod == 3) // с даты по дату с разбивкой по сменам
            {
                using (var dbContext = new ShiftReportDbContext())
                {
                    var result = from pqr in dbContext.ProdQualityReport
                                 join pc in dbContext.ProductCategories on pqr.Product equals pc.ProductID
                                 join sr in dbContext.ShiftReport on pqr.Report equals sr.ShiftReportID
                                 join dr in dbContext.ProdDefectReport on pqr.PQReportID equals dr.ProductReport into defectReports
                                 from dr in defectReports.DefaultIfEmpty() // Выполняем LEFT JOIN
                                 where sr.ShiftDate >= varDate1 && sr.ShiftDate <= varDate2 && shiftNumbs.Contains(sr.ShiftNum)
                                 group new { pqr, pc, dr } by new
                                 {
                                     sr.ShiftNum,
                                     pc.ProductName,
                                     pqr.Unspecified,
                                     pqr.ProdDepth,
                                     pqr.ProdLength,
                                     pqr.ProdWidth
                                 } into grp
                                 orderby grp.Key.ShiftNum
                                 select new
                                 {
                                     ProductName = grp.Key.ProductName,
                                     Unspecified = grp.Key.Unspecified,
                                     ProdDepth = grp.Key.ProdDepth,
                                     Length = grp.Key.ProdLength,
                                     Width = grp.Key.ProdWidth,
                                     ShiftNum = grp.Key.ShiftNum,
                                     AvgVolumePack = grp.Average(x => x.pqr.VolumePack),
                                     AvgDensityAvg = grp.Average(x => x.pqr.AvgDensity),
                                     PackCountSum = grp.Sum(x => x.pqr.PackCount),
                                     VolumeProdSum = grp.Sum(x => x.pqr.VolumeProduct),
                                     WeightSum = grp.Sum(x => x.pqr.Weight),
                                     AvgDensityOS = grp.Average(x => x.dr.DefectDensity),
                                     LowQualCount = grp.Sum(x => x.dr.DefectType >= 1 && x.dr.DefectType <= 7 ? x.dr.DefectPackCount : 0),
                                     LowQualVol = grp.Sum(x => x.dr.DefectType >= 1 && x.dr.DefectType <= 7 ? x.dr.DefectVolume : 0),
                                     LowQualWeight = grp.Sum(x => x.dr.DefectType >= 1 && x.dr.DefectType <= 7 ? x.dr.DefectWeight : 0),
                                     RejectVol = grp.Sum(x => x.dr.DefectType == 8 ? x.dr.DefectVolume : 0),
                                     RejectWeight = grp.Sum(x => x.dr.DefectType == 8 ? x.dr.DefectWeight : 0),
                                     PercentLowQual = (grp.Sum(x => x.dr.DefectType >= 1 && x.dr.DefectType <= 7 ? x.dr.DefectWeight : 0) / grp.Sum(x => x.pqr.Weight)) * 100.0,
                                     PercentReject = (grp.Sum(x => x.dr.DefectType == 8 ? x.dr.DefectWeight : 0) / grp.Sum(x => x.pqr.Weight)) * 100.0
                                 };
                    result.ToList();

                    DataTable dataTable = new DataTable();
                    dataTable.Columns.Add("Марка", typeof(string));
                    dataTable.Columns.Add("Неуказанная плт-ть", typeof(bool));
                    dataTable.Columns.Add("Толщина", typeof(int));
                    dataTable.Columns.Add("Длинна", typeof(int));
                    dataTable.Columns.Add("Ширина", typeof(int));
                    dataTable.Columns.Add("Смена", typeof(int));
                    dataTable.Columns.Add("Объем пачки", typeof(float));
                    dataTable.Columns.Add("Ср. плотность", typeof(float));
                    dataTable.Columns.Add("Кол-во пачек", typeof(int));
                    dataTable.Columns.Add("Объем", typeof(float));
                    dataTable.Columns.Add("Вес", typeof(float));
                    dataTable.Columns.Add("Средняя плотность ОС", typeof(float));
                    dataTable.Columns.Add("Кол-во пачек ОС", typeof(int));
                    dataTable.Columns.Add("Объем ОС", typeof(float));
                    dataTable.Columns.Add("Вес ОС", typeof(float));
                    dataTable.Columns.Add("Объем обрезь", typeof(float));
                    dataTable.Columns.Add("Вес обрезь", typeof(float));
                    dataTable.Columns.Add("Процент ОС", typeof(double));
                    dataTable.Columns.Add("Процент обрезь", typeof(double));

                    // Заполняем DataTable
                    foreach (var report in result)
                    {
                        DataRow row = dataTable.NewRow();
                        row["Марка"] = report.ProductName;
                        row["Неуказанная плт-ть"] = report.Unspecified;
                        row["Толщина"] = report.ProdDepth;
                        row["Длинна"] = report.Length;
                        row["Ширина"] = report.Width;
                        row["Смена"] = report.ShiftNum;
                        row["Объем пачки"] = report.AvgVolumePack;
                        row["Ср. плотность"] = report.AvgDensityAvg;
                        row["Кол-во пачек"] = report.PackCountSum;
                        row["Объем"] = report.VolumeProdSum;
                        row["Вес"] = report.WeightSum;
                        row["Средняя плотность ОС"] = report.AvgDensityOS;
                        row["Кол-во пачек ОС"] = report.LowQualCount;
                        row["Объем ОС"] = report.LowQualVol;
                        row["Вес ОС"] = report.LowQualWeight;
                        row["Объем обрезь"] = report.RejectVol;
                        row["Вес обрезь"] = report.RejectWeight;
                        row["Процент ОС"] = report.PercentLowQual;
                        row["Процент обрезь"] = report.PercentReject;
                        // Заполняйте остальные поля аналогично
                        dataTable.Rows.Add(row);
                    }
                    dbContext.Dispose();
                    return dataTable;
                }
            }
            //---------------------------------------------------------------------------------------------------------------------------------------
            if (getMethod == 4) // отчет с даты по дату по типам брака
            {
                using (var dbContext = new ShiftReportDbContext())
                {
                    var result = (
                        from pqr in dbContext.ProdQualityReport
                        join sr in dbContext.ShiftReport on pqr.Report equals sr.ShiftReportID
                        join dr in dbContext.ProdDefectReport on pqr.PQReportID equals dr.ProductReport into defectReports
                        from dr in defectReports.DefaultIfEmpty() // Выполняем LEFT JOIN
                        join dt in dbContext.DefectTypes on dr.DefectType equals dt.DefectTypeID
                        where sr.ShiftDate >= varDate1 && sr.ShiftDate <= varDate2
                        group new { dr, dt, sr, pqr }
                        by new 
                            {
                            dr.DefectType,
                            dt.DefectName 
                            }
                        into grp
                        orderby grp.Key.DefectType
                    select new
                    {
                        DefectTypes = grp.Key.DefectType,
                        DefectName = grp.Key.DefectName,
                        DefectPackCount = grp.Sum(x => x.dr.DefectPackCount),
                        DefectVolume = grp.Sum(x => x.dr.DefectVolume),
                        DefectWeight = grp.Sum(x => x.dr.DefectWeight)
                    }).ToList();


                    DataTable dataTable = new DataTable();
                    dataTable.Columns.Add("Тип дефекта", typeof(int));
                    dataTable.Columns.Add("Название дефекта", typeof(string));
                    dataTable.Columns.Add("Кол-во пачек", typeof(int));
                    dataTable.Columns.Add("Объем", typeof(float));
                    dataTable.Columns.Add("Вес", typeof(float));

                    // Заполняем DataTable
                    foreach (var report in result)
                    {
                        DataRow row = dataTable.NewRow();
                        row["Тип дефекта"] = report.DefectTypes;
                        row["Название дефекта"] = report.DefectName;
                        row["Кол-во пачек"] = report.DefectPackCount;
                        row["Объем"] = report.DefectVolume;
                        row["Вес"] = report.DefectWeight;
                        // Заполняйте остальные поля аналогично
                        dataTable.Rows.Add(row);
                    }
                    dbContext.Dispose();
                    return dataTable;
                }
            }
            //---------------------------------------------------------------------------------------------------------------------------------------
            if (getMethod == 5) // отчет с даты по дату по типам брака с разбивкой по номерам смен 
            {
                using (var dbContext = new ShiftReportDbContext())
                {
                    var result = (
                        from pqr in dbContext.ProdQualityReport
                        join sr in dbContext.ShiftReport on pqr.Report equals sr.ShiftReportID
                        join dr in dbContext.ProdDefectReport on pqr.PQReportID equals dr.ProductReport into defectReports
                        from dr in defectReports.DefaultIfEmpty() // Выполняем LEFT JOIN
                        join dt in dbContext.DefectTypes on dr.DefectType equals dt.DefectTypeID
                        where sr.ShiftDate >= varDate1 && sr.ShiftDate <= varDate2 && shiftNumbs.Contains(sr.ShiftNum)
                        group new { dr, dt, sr, pqr }
                        by new 
                        {
                            sr.ShiftNum,
                            dr.DefectType,
                            dt.DefectName 
                        } 
                        into grp
                        orderby grp.Key.ShiftNum
                        select new
                        {
                            ShiftNumer = grp.Key.ShiftNum,
                            DefectType = grp.Key.DefectType,
                            DefectName = grp.Key.DefectName,
                            DefectPackCount = grp.Sum(x => x.dr.DefectPackCount),
                            DefectVolume = grp.Sum(x => x.dr.DefectVolume),
                            DefectWeight = grp.Sum(x => x.dr.DefectWeight)
                        }).ToList();


                    DataTable dataTable = new DataTable();
                    dataTable.Columns.Add("Смена", typeof(int));
                    dataTable.Columns.Add("Тип дефекта", typeof(int));
                    dataTable.Columns.Add("Название дефекта", typeof(string));
                    dataTable.Columns.Add("Кол-во пачек", typeof(int));
                    dataTable.Columns.Add("Объем", typeof(float));
                    dataTable.Columns.Add("Вес", typeof(float));

                    // Заполняем DataTable
                    foreach (var report in result)
                    {
                        DataRow row = dataTable.NewRow();
                        row["Смена"] = report.ShiftNumer;
                        row["Тип дефекта"] = report.DefectType;
                        row["Название дефекта"] = report.DefectName;
                        row["Кол-во пачек"] = report.DefectPackCount;
                        row["Объем"] = report.DefectVolume;
                        row["Вес"] = report.DefectWeight;
                        // Заполняйте остальные поля аналогично
                        dataTable.Rows.Add(row);
                    }
                    dbContext.Dispose();
                    return dataTable;
                }
            }
            //---------------------------------------------------------------------------------------------------------------------------------------
            if (getMethod == 10) // отчет с даты по дату по простоям
            {
                using (var dbContext = new ShiftReportDbContext())
                {
                    var result = (
                        from str in dbContext.StopsReport
                        join sr in dbContext.ShiftReport on str.ShiftReport equals sr.ShiftReportID
                        join st in dbContext.StopType on str.StopType equals st.StopTypeID
                        join pil in dbContext.PlaceInLine on str.PlaceStop equals pil.PlacesID
                        where sr.ShiftDate >= varDate1 && sr.ShiftDate <= varDate2 && shiftDays.Contains(sr.ShiftDay)
                        group new { st, pil, str, sr } by new { st.StopCategory, st.StopName, pil.PlacesName } into grp
                        orderby grp.Sum(x => x.str.DurationStopMin)
                        select new
                        {
                            StopCategorys = grp.Key.StopCategory,
                            StopNames = grp.Key.StopName,
                            PlaceNames = grp.Key.PlacesName,
                            DurationStops = grp.Sum(x => x.str.DurationStopMin)
                            
                        }).ToList();


                    DataTable dataTable = new DataTable();
                    dataTable.Columns.Add("Категория остановки", typeof(string));
                    dataTable.Columns.Add("Название остановки", typeof(string));
                    dataTable.Columns.Add("Место остановки", typeof(string));
                    dataTable.Columns.Add("Длительность", typeof(float));


                    // Заполняем DataTable
                    foreach (var report in result)
                    {
                        DataRow row = dataTable.NewRow();
                        row["Категория остановки"] = report.StopCategorys;
                        row["Название остановки"] = report.StopNames;
                        row["Место остановки"] = report.PlaceNames;
                        row["Длительность"] = report.DurationStops;
                        // Заполняйте остальные поля аналогично
                        dataTable.Rows.Add(row);
                    }
                    dbContext.Dispose();
                    return dataTable;
                }
            }
            //---------------------------------------------------------------------------------------------------------------------------------------
            if (getMethod == 11) // отчет с даты по дату по простоям с разбивкой по видам простоя
            {
                using (var dbContext = new ShiftReportDbContext())
                {
                    var result = (
                        from str in dbContext.StopsReport
                        join sr in dbContext.ShiftReport on str.ShiftReport equals sr.ShiftReportID
                        join st in dbContext.StopType on str.StopType equals st.StopTypeID
                        join pil in dbContext.PlaceInLine on str.PlaceStop equals pil.PlacesID
                        where sr.ShiftDate >= varDate1 && sr.ShiftDate <= varDate2
                        group new { st, pil, str, sr } by new { st.StopCategory } into grp
                        orderby grp.Sum(x => x.str.DurationStopMin)
                        select new
                        {
                            StopCategorys = grp.Key.StopCategory,
                            DurationStops = grp.Sum(x => x.str.DurationStopMin)

                        }).ToList();


                    DataTable dataTable = new DataTable();
                    dataTable.Columns.Add("Категория остановки", typeof(string));
                    dataTable.Columns.Add("Длительность", typeof(float));


                    // Заполняем DataTable
                    foreach (var report in result)
                    {
                        DataRow row = dataTable.NewRow();
                        row["Категория остановки"] = report.StopCategorys;
                        row["Длительность"] = report.DurationStops;
                        // Заполняйте остальные поля аналогично
                        dataTable.Rows.Add(row);
                    }
                    dbContext.Dispose();
                    return dataTable;
                }
            }
            //---------------------------------------------------------------------------------------------------------------------------------------
            if (getMethod == 12) // отчет с даты по дату по простоям с разбивкой по номерам смен
            {
                using (var dbContext = new ShiftReportDbContext())
                {
                    var result = (
                        from str in dbContext.StopsReport
                        join sr in dbContext.ShiftReport on str.ShiftReport equals sr.ShiftReportID
                        join st in dbContext.StopType on str.StopType equals st.StopTypeID
                        join pil in dbContext.PlaceInLine on str.PlaceStop equals pil.PlacesID
                        where sr.ShiftDate >= varDate1 && sr.ShiftDate <= varDate2 && shiftNumbs.Contains(sr.ShiftNum)
                        group new { st, pil, str, sr } by new {sr.ShiftNum, st.StopCategory, st.StopName, pil.PlacesName } into grp
                        orderby grp.Key.ShiftNum, grp.Sum(x => x.str.DurationStopMin) descending
                        select new
                        {
                            ShiftNumer = grp.Key.ShiftNum,
                            StopCategorys = grp.Key.StopCategory,
                            StopNames = grp.Key.StopName,
                            PlaceNames = grp.Key.PlacesName,
                            DurationStops = grp.Sum(x => x.str.DurationStopMin)

                        }).ToList();


                    DataTable dataTable = new DataTable();
                    dataTable.Columns.Add("Номер смены", typeof(int));
                    dataTable.Columns.Add("Категория остановки", typeof(string));
                    dataTable.Columns.Add("Название остановки", typeof(string));
                    dataTable.Columns.Add("Место остановки", typeof(string));
                    dataTable.Columns.Add("Длительность", typeof(float));


                    // Заполняем DataTable
                    foreach (var report in result)
                    {
                        DataRow row = dataTable.NewRow();
                        row["Номер смены"] = report.ShiftNumer;
                        row["Категория остановки"] = report.StopCategorys;
                        row["Название остановки"] = report.StopNames;
                        row["Место остановки"] = report.PlaceNames;
                        row["Длительность"] = report.DurationStops;
                        // Заполняйте остальные поля аналогично
                        dataTable.Rows.Add(row);
                    }
                    dbContext.Dispose();
                    return dataTable;
                }
            }
            //---------------------------------------------------------------------------------------------------------------------------------------
            if (getMethod == 13) // отчет с даты по дату с комментариями по простоям
            {
                using (var dbContext = new ShiftReportDbContext())
                {
                    var result = (
                        from str in dbContext.StopsReport
                        join sr in dbContext.ShiftReport on str.ShiftReport equals sr.ShiftReportID
                        join st in dbContext.StopType on str.StopType equals st.StopTypeID
                        join pil in dbContext.PlaceInLine on str.PlaceStop equals pil.PlacesID
                        where sr.ShiftDate >= varDate1 && sr.ShiftDate <= varDate2
                        group new { st, pil, str, sr } by new {sr.ShiftDate, st.StopCategory, st.StopName, pil.PlacesName, str.DurationStopMin, str.CommentStop } into grp
                        orderby grp.Key.ShiftDate
                        select new
                        {
                            ShiftDates = grp.Key.ShiftDate,
                            StopCategorys = grp.Key.StopCategory,
                            StopNames = grp.Key.StopName,
                            PlaceNames = grp.Key.PlacesName,
                            DurationStops = grp.Key.DurationStopMin,
                            StopComment = grp.Key.CommentStop

                        }).ToList();


                    DataTable dataTable = new DataTable();
                    dataTable.Columns.Add("Дата", typeof(DateTime));
                    dataTable.Columns.Add("Категория остановки", typeof(string));
                    dataTable.Columns.Add("Название остановки", typeof(string));
                    dataTable.Columns.Add("Место остановки", typeof(string));
                    dataTable.Columns.Add("Длительность", typeof(float));
                    dataTable.Columns.Add("Комментарий", typeof(string));


                    // Заполняем DataTable
                    foreach (var report in result)
                    {
                        DataRow row = dataTable.NewRow();
                        row["Дата"] = report.ShiftDates;
                        row["Категория остановки"] = report.StopCategorys;
                        row["Название остановки"] = report.StopNames;
                        row["Место остановки"] = report.PlaceNames;
                        row["Длительность"] = report.DurationStops;
                        row["Комментарий"] = report.StopComment;
                        // Заполняйте остальные поля аналогично
                        dataTable.Rows.Add(row);
                    }
                    dbContext.Dispose();
                    return dataTable;
                }
            }
            //---------------------------------------------------------------------------------------------------------------------------------------
            if (getMethod == 14) // отчет с даты по дату по простоям с разбивкой по видам простоя
            {
                using (var dbContext = new ShiftReportDbContext())
                {
                    var result = (
                        from str in dbContext.StopsReport
                        join sr in dbContext.ShiftReport on str.ShiftReport equals sr.ShiftReportID
                        join st in dbContext.StopType on str.StopType equals st.StopTypeID
                        join pil in dbContext.PlaceInLine on str.PlaceStop equals pil.PlacesID
                        where sr.ShiftDate >= varDate1 && sr.ShiftDate <= varDate2 && shiftDays.Contains(sr.ShiftNum) && stopCategoryes.Contains(st.StopCategory)
                        group new { st, pil, str, sr } by new { sr.ShiftNum, st.StopCategory, st.StopName, pil.PlacesName } into grp
                        orderby grp.Key.ShiftNum, grp.Sum(x => x.str.DurationStopMin) descending
                        select new
                        {
                            ShiftNumer = grp.Key.ShiftNum,
                            StopCategorys = grp.Key.StopCategory,
                            PlaceNames = grp.Key.PlacesName,
                            DurationStops = grp.Sum(x => x.str.DurationStopMin)

                        }).ToList();


                    DataTable dataTable = new DataTable();
                    dataTable.Columns.Add("Номер смены", typeof(int));
                    dataTable.Columns.Add("Категория остановки", typeof(string));
                    dataTable.Columns.Add("Место остановки", typeof(string));
                    dataTable.Columns.Add("Длительность", typeof(float));


                    // Заполняем DataTable
                    foreach (var report in result)
                    {
                        DataRow row = dataTable.NewRow();
                        row["Номер смены"] = report.ShiftNumer;
                        row["Категория остановки"] = report.StopCategorys;
                        row["Место остановки"] = report.PlaceNames;
                        row["Длительность"] = report.DurationStops;
                        // Заполняйте остальные поля аналогично
                        dataTable.Rows.Add(row);
                    }
                    dbContext.Dispose();
                    return dataTable;
                }
            }
            //---------------------------------------------------------------------------------------------------------------------------------------
            if (getMethod == 15) // отчет с даты по дату по простоям
            {
                using (var dbContext = new ShiftReportDbContext())
                {
                    var result = (
                        from str in dbContext.StopsReport
                        join sr in dbContext.ShiftReport on str.ShiftReport equals sr.ShiftReportID
                        join st in dbContext.StopType on str.StopType equals st.StopTypeID
                        join pil in dbContext.PlaceInLine on str.PlaceStop equals pil.PlacesID
                        where sr.ShiftDate >= varDate1 && sr.ShiftDate <= varDate2 && shiftDays.Contains(sr.ShiftDay)
                        group new { st, pil, str, sr }
                        by new
                        {
                            sr.ShiftDate,
                            sr.ShiftDay,
                            sr.ShiftNum,
                            st.StopCategory,
                            st.StopName,
                            pil.PlacesName,
                            pil.Section,
                            str.StopFirstTime,
                            str.StopEndTime,
                            str.DurationStopMin,
                            str.BreakdownWithoutStop,
                            str.CommentStop,
                            str.Centrifuge
                        } into grp
                        orderby grp.Key.StopFirstTime
                        select new
                        {
                            ShiftNumer = grp.Key.ShiftNum,
                            PlaceNames = grp.Key.PlacesName,
                            Section = grp.Key.Section,
                            StopCategorys = grp.Key.StopCategory,
                            StopNames = grp.Key.StopName,
                            FirstTime = grp.Key.StopFirstTime,
                            EndTime = grp.Key.StopEndTime,
                            DurationStops = grp.Key.DurationStopMin,
                            BreakdownWithoutStop = grp.Key.BreakdownWithoutStop,
                            Comment = grp.Key.CommentStop,
                            DateChangeFuge = grp.Key.ShiftDate,
                            Fuge = grp.Key.Centrifuge
                        });
                    result.ToList();

                    DataTable dataTable = new DataTable();
                    dataTable.Columns.Add("Номер смены", typeof(int));
                    dataTable.Columns.Add("Место остановки", typeof(string));
                    dataTable.Columns.Add("Узел", typeof(string));
                    dataTable.Columns.Add("Категория остановки", typeof(string));
                    dataTable.Columns.Add("Название остановки", typeof(string));
                    dataTable.Columns.Add("Начало остановки", typeof(string));
                    dataTable.Columns.Add("Конец остановки", typeof(string));
                    dataTable.Columns.Add("Длительность", typeof(float));
                    dataTable.Columns.Add("Остановка выпуска", typeof(bool));
                    dataTable.Columns.Add("Комментарий", typeof(string));
                    dataTable.Columns.Add("Дата замены фуги", typeof(string));
                    dataTable.Columns.Add("Фуга", typeof(int));

                    // Заполняем DataTable
                    foreach (var report2 in result)
                    {
                        DataRow row = dataTable.NewRow();
                        row["Номер смены"] = report2.ShiftNumer;
                        row["Место остановки"] = report2.PlaceNames;
                        row["Узел"] = report2.Section;
                        row["Категория остановки"] = report2.StopCategorys;
                        row["Название остановки"] = report2.StopNames;
                        row["Начало остановки"] = report2.FirstTime;
                        row["Конец остановки"] = report2.EndTime;
                        row["Длительность"] = report2.DurationStops;
                        row["Остановка выпуска"] = report2.BreakdownWithoutStop;
                        row["Комментарий"] = report2.Comment;
                        row["Дата замены фуги"] = $"{report2.DateChangeFuge.Date} {report2.EndTime}";
                        row["Фуга"] = report2.Fuge;
                        // Заполняйте остальные поля аналогично
                        dataTable.Rows.Add(row);
                    }
                    dbContext.Dispose();
                    return dataTable;
                }
            }
            //---------------------------------------------------------------------------------------------------------------------------------------
            if (getMethod == 30) // отчет по продуктам
            {
                using (var dbContext = new ShiftReportDbContext())
                {
                    var result = from pqr in dbContext.ProdQualityReport
                    join sr in dbContext.ShiftReport on pqr.Report equals sr.ShiftReportID
                    join pc in dbContext.ProductCategories on pqr.Product equals pc.ProductID
                    where sr.ShiftDate >= varDate1.Date && sr.ShiftDate <= varDate2.Date && shiftDays.Contains(sr.ShiftDay)
                    group new { pqr, sr } by new
                    {
                        sr.ShiftReportID,
                        sr.ShiftDate,
                        sr.ShiftNum,
                        sr.ShiftDay,
                        pqr.PQReportID,
                        pc.ProductName,
                        pqr.Report,
                        pqr.Unspecified,
                        pqr.Regarding,
                        pqr.ProdLength,
                        pqr.ProdWidth,
                        pqr.ProdDepth,
                        pqr.VolumePack,
                        pqr.AvgDensity,
                        pqr.PackCount,
                        pqr.Weight,
                        pqr.VolumeProduct

                    } into grp
                    orderby grp.Key.ShiftReportID
                    select new
                    {
                        ShiftDate = grp.Key.ShiftDate,
                        ShiftNum = grp.Key.ShiftNum,
                        ShiftDay = grp.Key.ShiftDay,
                        NumReport = grp.Key.Report,
                        ProdReportID = grp.Key.PQReportID,
                        ProductName = grp.Key.ProductName,
                        Length = grp.Key.ProdLength,
                        Width = grp.Key.ProdWidth,
                        ProdDepth = grp.Key.ProdDepth,
                        Unspecified = grp.Key.Unspecified,
                        Regarding = grp.Key.Regarding,
                        AvgDensity = grp.Key.AvgDensity,
                        VolumePack = grp.Key.VolumePack,
                        PackCount = grp.Key.PackCount,
                        Weight = grp.Key.Weight,
                        VolumeProduct = grp.Key.VolumeProduct
                    };
                    result.ToList();

                    DataTable dataTable = new DataTable();
                    dataTable.Columns.Add("Дата", typeof(DateTime));
                    dataTable.Columns.Add("Смена", typeof(int));
                    dataTable.Columns.Add("День/Ночь", typeof(int));
                    dataTable.Columns.Add("# записи смены", typeof(int));
                    dataTable.Columns.Add("# записи", typeof(int));
                    dataTable.Columns.Add("Марка", typeof(string));
                    dataTable.Columns.Add("Длинна", typeof(int));
                    dataTable.Columns.Add("Ширина", typeof(int));
                    dataTable.Columns.Add("Толщина", typeof(int));
                    dataTable.Columns.Add("Неуказанная плт-ть", typeof(bool));
                    dataTable.Columns.Add("Пересорт", typeof(bool));
                    dataTable.Columns.Add("Ср. плотность", typeof(float));
                    dataTable.Columns.Add("Объем пачки", typeof(float));
                    dataTable.Columns.Add("Кол-во пачек", typeof(int));
                    dataTable.Columns.Add("Объем", typeof(float));
                    dataTable.Columns.Add("Вес", typeof(float));
                // Заполняем DataTable
                    foreach (var report in result)
                    {
                        DataRow row = dataTable.NewRow();
                        row["Дата"] = report.ShiftDate;
                        row["Смена"] = report.ShiftNum;
                        row["День/Ночь"] = report.ShiftDay;
                        row["# записи смены"] = report.NumReport;
                        row["# записи"] = report.ProdReportID;
                        row["Марка"] = report.ProductName;
                        row["Длинна"] = report.Length;
                        row["Ширина"] = report.Width;
                        row["Толщина"] = report.ProdDepth;
                        row["Неуказанная плт-ть"] = report.Unspecified;
                        row["Пересорт"] = report.Regarding;
                        row["Ср. плотность"] = report.AvgDensity;
                        row["Объем пачки"] = report.VolumePack;
                        row["Кол-во пачек"] = report.PackCount;
                        row["Объем"] = report.VolumeProduct;
                        row["Вес"] = report.Weight;
                        // Заполняйте остальные поля аналогично
                        dataTable.Rows.Add(row);
                    }
                    dbContext.Dispose();
                    return dataTable;
                }
            }
            
            //---------------------------------------------------------------------------------------------------------------------------------------
            if (getMethod == 31) // отчет по дефектам
            {
                using (var dbContext = new ShiftReportDbContext())
                {
                    var result = from pdr in dbContext.ProdDefectReport
                        join pqr in dbContext.ProdQualityReport on pdr.ProductReport equals pqr.PQReportID
                        join sr in dbContext.ShiftReport on pqr.Report equals sr.ShiftReportID
                        join pc in dbContext.ProductCategories on pqr.Product equals pc.ProductID
                        join dt in dbContext.DefectTypes on pdr.DefectType equals dt.DefectTypeID
                        where sr.ShiftDate >= varDate1.Date && sr.ShiftDate <= varDate2.Date && shiftDays.Contains(sr.ShiftDay)
                        group new { pqr, pdr, pc, sr } by new
                        {
                            sr.ShiftReportID,
                            sr.ShiftDate,
                            sr.ShiftNum,
                            sr.ShiftDay,
                            pqr.PQReportID,
                            pdr.DefectReportID,
                            pc.ProductName,
                            dt.DefectName,
                            pdr.DefectType,
                            pdr.DefectVolumePack,
                            pdr.DefectDensity,
                            pdr.DefectPackCount,
                            pdr.DefectVolume,
                            pdr.DefectWeight,

                        } into grp
                        orderby grp.Key.ShiftReportID
                        select new
                        {
                            ShiftDate = grp.Key.ShiftDate,
                            ShiftNum = grp.Key.ShiftNum,
                            ShiftDay = grp.Key.ShiftDay,
                            NumReport = grp.Key.PQReportID,
                            DefectReportID = grp.Key.DefectReportID,
                            ProductName = grp.Key.ProductName,
                            DefectName = grp.Key.DefectName,
                            DefectType = grp.Key.DefectType,
                            DefectVolumePack = grp.Key.DefectVolumePack,
                            DefectDensity = grp.Key.DefectDensity,
                            DefectPackCount = grp.Key.DefectPackCount,
                            DefectVolume = grp.Key.DefectVolume,
                            DefectWeight = grp.Key.DefectWeight
                        };
                    result.ToList();

                    DataTable dataTable = new DataTable();
                    dataTable.Columns.Add("Дата", typeof(DateTime));
                    dataTable.Columns.Add("Смена", typeof(int));
                    dataTable.Columns.Add("День/Ночь", typeof(int));
                    dataTable.Columns.Add("# записи продукта", typeof(int));
                    dataTable.Columns.Add("# записи", typeof(int));
                    dataTable.Columns.Add("Марка", typeof(string));
                    dataTable.Columns.Add("Дефект", typeof(string));
                    dataTable.Columns.Add("Тип дефекта", typeof(string));
                    dataTable.Columns.Add("Объем пачки", typeof(float));
                    dataTable.Columns.Add("Плотность", typeof(float));
                    dataTable.Columns.Add("Кол-во пачек", typeof(int));
                    dataTable.Columns.Add("Объем", typeof(float));
                    dataTable.Columns.Add("Вес", typeof(float));
                    // Заполняем DataTable
                    foreach (var report in result)
                    {
                        DataRow row = dataTable.NewRow();
                        row["Дата"] = report.ShiftDate;
                        row["Смена"] = report.ShiftNum;
                        row["День/Ночь"] = report.ShiftDay;
                        row["# записи продукта"] = report.NumReport;
                        row["# записи"] = report.DefectReportID;
                        row["Марка"] = report.ProductName;
                        row["Дефект"] = report.DefectName;
                        row["Тип дефекта"] = report.DefectType;
                        row["Объем пачки"] = report.DefectVolumePack;
                        row["Плотность"] = report.DefectDensity;
                        row["Кол-во пачек"] = report.DefectPackCount;
                        row["Объем"] = report.DefectVolume;
                        row["Вес"] = report.DefectWeight;
                        // Заполняйте остальные поля аналогично
                        dataTable.Rows.Add(row);
                    }
                    dbContext.Dispose();
                    return dataTable;
                }
            }
            //---------------------------------------------------------------------------------------------------------------------------------------
            if (getMethod == 32) // отчет технический
            {
                using (var dbContext = new ShiftReportDbContext())
                {
                    var result = from str in dbContext.StopsReport
                        join sr in dbContext.ShiftReport on str.ShiftReport equals sr.ShiftReportID
                        join pil in dbContext.PlaceInLine on str.PlaceStop equals pil.PlacesID
                        join st in dbContext.StopType on str.StopType equals st.StopTypeID
                        where sr.ShiftDate >= varDate1.Date && sr.ShiftDate <= varDate2.Date && shiftDays.Contains(sr.ShiftDay)
                        group new { str, pil, st, sr } by new
                        {
                            sr.ShiftReportID,
                            sr.ShiftDate,
                            sr.ShiftNum,
                            sr.ShiftDay,
                            str.ShiftReport,
                            str.StopReportID,
                            st.StopCategory,
                            st.StopName,
                            pil.PlacesName,
                            pil.Section,
                            str.StopFirstTime,
                            str.StopEndTime,
                            str.DurationStopMin,
                            str.BreakdownWithoutStop,
                            str.CommentStop,
                             str.Centrifuge

                        } into grp
                    orderby grp.Key.ShiftReportID
                    select new
                    {
                        ShiftDate = grp.Key.ShiftDate,
                        ShiftNum = grp.Key.ShiftNum,
                        ShiftDay = grp.Key.ShiftDay,
                        NumReport = grp.Key.ShiftReport,
                        StopReportID = grp.Key.StopReportID,
                        StopCategorys = grp.Key.StopCategory,
                        StopNames = grp.Key.StopName,
                        PlaceNames = grp.Key.PlacesName,
                        Section = grp.Key.Section,
                        FirstTime = grp.Key.StopFirstTime,
                        EndTime = grp.Key.StopEndTime,
                        DurationStops = grp.Key.DurationStopMin,
                        BreakdownWithoutStop = grp.Key.BreakdownWithoutStop,
                        Comment = grp.Key.CommentStop,
                        Fuge = grp.Key.Centrifuge
                    };
                    result.ToList();

                    DataTable dataTable = new DataTable();
                    dataTable.Columns.Add("Дата", typeof(DateTime));
                    dataTable.Columns.Add("Смена", typeof(int));
                    dataTable.Columns.Add("День/Ночь", typeof(int));
                    dataTable.Columns.Add("# записи отчета", typeof(int));
                    dataTable.Columns.Add("# записи", typeof(int));
                    dataTable.Columns.Add("Категория остановки", typeof(string));
                    dataTable.Columns.Add("Название остановки", typeof(string));
                    dataTable.Columns.Add("Место остановки", typeof(string));
                    dataTable.Columns.Add("Узел", typeof(string));
                    dataTable.Columns.Add("Начало остановки", typeof(string));
                    dataTable.Columns.Add("Конец остановки", typeof(string));
                    dataTable.Columns.Add("Длительность", typeof(float));
                    dataTable.Columns.Add("Остановка выпуска", typeof(bool));
                    dataTable.Columns.Add("Комментарий", typeof(string));
                    dataTable.Columns.Add("Фуга", typeof(int));
                    // Заполняем DataTable
                    foreach (var report in result)
                    {
                        DataRow row = dataTable.NewRow();
                        row["Дата"] = report.ShiftDate;
                        row["Смена"] = report.ShiftNum;
                        row["День/Ночь"] = report.ShiftDay;
                        row["# записи отчета"] = report.NumReport;
                        row["# записи"] = report.StopReportID;
                        row["Категория остановки"] = report.StopCategorys;
                        row["Название остановки"] = report.StopNames;
                        row["Место остановки"] = report.PlaceNames;
                        row["Узел"] = report.Section;
                        row["Начало остановки"] = report.FirstTime;
                        row["Конец остановки"] = report.EndTime;
                        row["Длительность"] = report.DurationStops;
                        row["Остановка выпуска"] = report.BreakdownWithoutStop;
                        row["Комментарий"] = report.Comment;
                        row["Фуга"] = report.Fuge;
                        // Заполняйте остальные поля аналогично
                        dataTable.Rows.Add(row);
                    }
                    dbContext.Dispose();
                    return dataTable;
                }
            }
            //---------------------------------------------------------------------------------------------------------------------------------------
            if (getMethod == 33) // продукты
            {
                using (var dbContext = new ShiftReportDbContext())
                {
                    var result = from pc in dbContext.ProductCategories

                    group new { pc } by new
                    {
                        pc.ProductID,
                        pc.ProductName,
                        pc.DensityMin,
                        pc.DensityMax
                    } into grp
                    orderby grp.Key.ProductID
                    select new
                    {
                        ProductID = grp.Key.ProductID,
                        ProdName = grp.Key.ProductName,
                        DensityMin = grp.Key.DensityMin,
                        DensityMax = grp.Key.DensityMax
                    };
                    result.ToList();

                    DataTable dataTable = new DataTable();
                    dataTable.Columns.Add("ID", typeof(int));
                    dataTable.Columns.Add("Имя продукта", typeof(string));
                    dataTable.Columns.Add("Минимальная плотность", typeof(float));
                    dataTable.Columns.Add("Максимальная плотность", typeof(float));
                    // Заполняем DataTable
                    foreach (var report in result)
                    {
                        DataRow row = dataTable.NewRow();
                        row["ID"] = report.ProductID;
                        row["Имя продукта"] = report.ProdName;
                        row["Минимальная плотность"] = report.DensityMin;
                        row["Максимальная плотность"] = report.DensityMax;
                        // Заполняйте остальные поля аналогично
                        dataTable.Rows.Add(row);
                    }
                    dbContext.Dispose();
                    return dataTable;
                }
            }
            //---------------------------------------------------------------------------------------------------------------------------------------
            if (getMethod == 34) // дефекты
            {
                using (var dbContext = new ShiftReportDbContext())
                {
                    var result = from dt in dbContext.DefectTypes

                    group new { dt } by new
                    {
                        dt.DefectTypeID,
                        dt.DefectName
                    } into grp
                    orderby grp.Key.DefectTypeID
                    select new
                    {
                        ID = grp.Key.DefectTypeID,
                        DefectName = grp.Key.DefectName
                    };
                    result.ToList();

                    DataTable dataTable = new DataTable();
                    dataTable.Columns.Add("ID", typeof(int));
                    dataTable.Columns.Add("Имя дефекта", typeof(string));
                    // Заполняем DataTable
                    foreach (var report in result)
                    {
                        DataRow row = dataTable.NewRow();
                        row["ID"] = report.ID;
                        row["Имя дефекта"] = report.DefectName;
                        // Заполняйте остальные поля аналогично
                        dataTable.Rows.Add(row);
                    }
                    dbContext.Dispose();
                    return dataTable;
                }
            }
            //---------------------------------------------------------------------------------------------------------------------------------------
            if (getMethod == 35) // остановки
            {
                using (var dbContext = new ShiftReportDbContext())
                {
                    var result = from st in dbContext.StopType

                    group new { st } by new
                    {
                    st.StopTypeID,
                    st.StopName,
                    st.StopCategory
                    } into grp
                    orderby grp.Key.StopTypeID
                    select new
                    {
                        ID = grp.Key.StopTypeID,
                        StopName = grp.Key.StopName,
                        StopCategory = grp.Key.StopCategory
                    };
                    result.ToList();

                    DataTable dataTable = new DataTable();
                    dataTable.Columns.Add("ID", typeof(string));
                    dataTable.Columns.Add("Остановка", typeof(string));
                    dataTable.Columns.Add("Категория", typeof(string));
                    // Заполняем DataTable
                    foreach (var report in result)
                    {
                        DataRow row = dataTable.NewRow();
                        row["ID"] = report.ID;
                        row["Остановка"] = report.StopName;
                        row["Категория"] = report.StopCategory;
                        // Заполняйте остальные поля аналогично
                        dataTable.Rows.Add(row);
                    }
                    dbContext.Dispose();
                    return dataTable;
                }   
            }
            //---------------------------------------------------------------------------------------------------------------------------------------
            if (getMethod == 36) // места на линии
            {
                using (var dbContext = new ShiftReportDbContext())
                {
                    var result = from pil in dbContext.PlaceInLine

                    group new { pil } by new
                    {
                        pil.PlacesID,
                        pil.PlacesName,
                        pil.Section
                    } into grp
                    orderby grp.Key.PlacesID
                    select new
                    {
                        ID = grp.Key.PlacesID,
                        PlacesName = grp.Key.PlacesName,
                        Section = grp.Key.Section
                    };
                    result.ToList();

                    DataTable dataTable = new DataTable();
                    dataTable.Columns.Add("ID", typeof(int));
                    dataTable.Columns.Add("Место", typeof(string));
                    dataTable.Columns.Add("Машина", typeof(string));
                    // Заполняем DataTable
                    foreach (var report in result)
                    {
                        DataRow row = dataTable.NewRow();
                        row["ID"] = report.ID;
                        row["Место"] = report.PlacesName;
                        row["Машина"] = report.Section;
                        // Заполняйте остальные поля аналогично
                        dataTable.Rows.Add(row);
                    }
                    dbContext.Dispose();
                    return dataTable;
                }
            }
            throw new Exception("Специальное условие прерывает метод");
        }
        //---------------------------------------------------------------------------------------------------------------------------------------

        public List <DataTable> ExtractProductList(int getMethod, DateTime varDate1, DateTime varDate2, List<int> shiftDays, List<int> shiftNumbs, List<string> stopCategoryes)
        {
            if (getMethod == 21) // отчет за предыдущие сутки с разбивкой по сменам
            {
                using (var dbContext = new ShiftReportDbContext())
                {
                    var result = from pqr in dbContext.ProdQualityReport
                                 join pc in dbContext.ProductCategories on pqr.Product equals pc.ProductID
                                 join sr in dbContext.ShiftReport on pqr.Report equals sr.ShiftReportID
                                 join pdr in dbContext.ProdDefectReport.DefaultIfEmpty() on pqr.PQReportID equals pdr.ProductReport
                                 where sr.ShiftDate >= varDate1.Date && sr.ShiftDate <= varDate2.Date && shiftDays.Contains(sr.ShiftDay)
                                 group new { pqr, pc, pdr } by new
                                 {
                                     sr.ShiftReportID,
                                     sr.ShiftDate,
                                     sr.ShiftNum,
                                     pqr.VolumePack,
                                     pc.ProductName,
                                     pqr.Unspecified,
                                     pqr.Regarding,
                                     pqr.ProdDepth,
                                     pqr.ProdLength,
                                     pqr.ProdWidth
                                 } into grp
                                 orderby grp.Key.ShiftReportID
                                 select new
                                 {
                                     ShiftDate = grp.Key.ShiftDate,
                                     ShiftNum = grp.Key.ShiftNum,
                                     ProductName = grp.Key.ProductName,
                                     Unspecified = grp.Key.Unspecified,
                                     ProdDepth = grp.Key.ProdDepth,
                                     Length = grp.Key.ProdLength,
                                     Width = grp.Key.ProdWidth,
                                     AvgDensityAvg = grp.Average(x => x.pqr.AvgDensity),
                                     VolumePack = grp.Key.VolumePack,
                                     PackCountSum = grp.Average(x => x.pqr.PackCount),
                                     VolumeProdSum = grp.Average(x => x.pqr.VolumeProduct),
                                     WeightSum = grp.Average(x => x.pqr.Weight),
                                     LowQualCount = grp.Sum(x => x.pdr.DefectType >= 1 && x.pdr.DefectType <= 7 ? x.pdr.DefectPackCount : 0),
                                     LowQualVol = grp.Sum(x => x.pdr.DefectType >= 1 && x.pdr.DefectType <= 7 ? x.pdr.DefectVolume : 0),
                                     LowQualWeight = grp.Sum(x => x.pdr.DefectType >= 1 && x.pdr.DefectType <= 7 ? x.pdr.DefectWeight : 0),
                                     RejectVol = grp.Sum(x => x.pdr.DefectType == 8 ? x.pdr.DefectVolume : 0),
                                     RejectWeight = grp.Sum(x => x.pdr.DefectType == 8 ? x.pdr.DefectWeight : 0),
                                     Regarding = grp.Key.Regarding,
                                     AvgLowQualDensity = grp.Average(x => x.pdr.DefectDensity)
                                 };
                    result.ToList();

                    var result2 = (
                        from str in dbContext.StopsReport
                        join sr in dbContext.ShiftReport on str.ShiftReport equals sr.ShiftReportID
                        join st in dbContext.StopType on str.StopType equals st.StopTypeID
                        join pil in dbContext.PlaceInLine on str.PlaceStop equals pil.PlacesID
                        where sr.ShiftDate >= varDate1 && sr.ShiftDate <= varDate2 && shiftDays.Contains(sr.ShiftDay)
                        group new { st, pil, str, sr }
                        by new 
                        { 
                            sr.ShiftDate,
                            sr.ShiftDay,
                            sr.ShiftNum,
                            st.StopCategory,
                            st.StopName,
                            pil.PlacesName,
                            pil.Section,
                            str.StopFirstTime,
                            str.StopEndTime,
                            str.DurationStopMin,
                            str.BreakdownWithoutStop,
                            str.CommentStop,
                            str.Centrifuge
                        } into grp
                        orderby grp.Key.StopFirstTime
                        select new
                        {
                            ShiftNumer = grp.Key.ShiftNum,
                            PlaceNames = grp.Key.PlacesName,
                            Section = grp.Key.Section,
                            StopCategorys = grp.Key.StopCategory,
                            StopNames = grp.Key.StopName,
                            FirstTime = grp.Key.StopFirstTime,
                            EndTime = grp.Key.StopEndTime,
                            DurationStops = grp.Key.DurationStopMin,
                            BreakdownWithoutStop = grp.Key.BreakdownWithoutStop,
                            Comment = grp.Key.CommentStop,
                            DateChangeFuge = grp.Key.ShiftDate,
                            Fuge = grp.Key.Centrifuge
                        });
                    result2.ToList();

                    DataTable dataTable = new DataTable();
                    dataTable.Columns.Add("Дата", typeof(DateTime));
                    dataTable.Columns.Add("Смена", typeof(int));
                    dataTable.Columns.Add("Марка", typeof(string));
                    dataTable.Columns.Add("Неуказанная плт-ть", typeof(bool));
                    dataTable.Columns.Add("Толщина", typeof(int));
                    dataTable.Columns.Add("Длинна", typeof(int));
                    dataTable.Columns.Add("Ширина", typeof(int));
                    dataTable.Columns.Add("Ср. плотность", typeof(float));
                    dataTable.Columns.Add("Объем пачки", typeof(float));
                    dataTable.Columns.Add("Кол-во пачек", typeof(int));
                    dataTable.Columns.Add("Объем", typeof(float));
                    dataTable.Columns.Add("Вес", typeof(float));
                    dataTable.Columns.Add("Кол-во пачек ОС", typeof(int));
                    dataTable.Columns.Add("Объем ОС", typeof(float));
                    dataTable.Columns.Add("Вес ОС", typeof(float));
                    dataTable.Columns.Add("Объем обрезь", typeof(float));
                    dataTable.Columns.Add("Вес обрезь", typeof(float));
                    dataTable.Columns.Add("Пересорт", typeof(bool));
                    dataTable.Columns.Add("Плотность ОС", typeof(float));

                    // Заполняем DataTable
                    foreach (var report in result)
                    {
                        DataRow row = dataTable.NewRow();
                        row["Дата"] = report.ShiftDate;
                        row["Смена"] = report.ShiftNum;
                        row["Марка"] = report.ProductName;
                        row["Неуказанная плт-ть"] = report.Unspecified;
                        row["Толщина"] = report.ProdDepth;
                        row["Длинна"] = report.Length;
                        row["Ширина"] = report.Width;
                        row["Ср. плотность"] = report.AvgDensityAvg;
                        row["Объем пачки"] = report.VolumePack;
                        row["Кол-во пачек"] = report.PackCountSum;
                        row["Объем"] = report.VolumeProdSum;
                        row["Вес"] = report.WeightSum;
                        row["Кол-во пачек ОС"] = report.LowQualCount;
                        row["Объем ОС"] = report.LowQualVol;
                        row["Вес ОС"] = report.LowQualWeight;
                        row["Объем обрезь"] = report.RejectVol;
                        row["Вес обрезь"] = report.RejectWeight;
                        row["Пересорт"] = report.Regarding;
                        row["Плотность ОС"] = report.AvgLowQualDensity;
                        // Заполняйте остальные поля аналогично
                        dataTable.Rows.Add(row);

                    }
                    DataTable dataTable2 = new DataTable();
                    dataTable2.Columns.Add("Номер смены", typeof(int));
                    dataTable2.Columns.Add("Место остановки", typeof(string));
                    dataTable2.Columns.Add("Узел", typeof(string));
                    dataTable2.Columns.Add("Категория остановки", typeof(string));
                    dataTable2.Columns.Add("Название остановки", typeof(string));
                    dataTable2.Columns.Add("Начало остановки", typeof(string));
                    dataTable2.Columns.Add("Конец остановки", typeof(string));
                    dataTable2.Columns.Add("Длительность", typeof(float));
                    dataTable2.Columns.Add("Остановка выпуска", typeof(bool));
                    dataTable2.Columns.Add("Комментарий", typeof(string));
                    dataTable2.Columns.Add("Дата замены фуги", typeof(string));
                    dataTable2.Columns.Add("Фуга", typeof(int));

                    // Заполняем DataTable
                    foreach (var report2 in result2)
                    {
                        DataRow row = dataTable2.NewRow();
                        row["Номер смены"] = report2.ShiftNumer;
                        row["Место остановки"] = report2.PlaceNames;
                        row["Узел"] = report2.Section;
                        row["Категория остановки"] = report2.StopCategorys;
                        row["Название остановки"] = report2.StopNames;
                        row["Начало остановки"] = report2.FirstTime;
                        row["Конец остановки"] = report2.EndTime;
                        row["Длительность"] = report2.DurationStops;
                        row["Остановка выпуска"] = report2.BreakdownWithoutStop;
                        row["Комментарий"] = report2.Comment;
                        row["Дата замены фуги"] = $"{report2.DateChangeFuge.Date} {report2.EndTime}";
                        row["Фуга"] = report2.Fuge;
                        // Заполняйте остальные поля аналогично
                        dataTable2.Rows.Add(row);
                    }
                    List<DataTable> comboTable = new List<DataTable> { dataTable, dataTable2 };
                    dbContext.Dispose();
                    return comboTable;
                }
            }
            throw new Exception("Специальное условие прерывает метод");
        }
        public List<ItemReportList> SetReportList(DateTime varDate1, DateTime varDate2, IEnumerable<int> shiftDays)
        {
            using (var dbContext = new ShiftReportDbContext())
            {
                var prodQualityReports = dbContext.ProdQualityReport
                    .Where(pqr =>
                        dbContext.ShiftReport
                            .Where(sr => sr.ShiftDate >= varDate1.Date && sr.ShiftDate <= varDate2.Date && shiftDays.Contains(sr.ShiftDay))
                            .Select(sr => sr.ShiftReportID)
                            .Contains(pqr.Report))
                    .Join(
                        dbContext.ProductCategories,
                        pqr => pqr.Product,
                        pc => pc.ProductID,
                        (pqr, pc) => new ItemReportList
                        {
                            ProductNames = pc.ProductName,
                            Unspecifies = pqr.Unspecified,
                            Depth = pqr.ProdDepth,
                            Length = pqr.ProdLength,
                            Width = pqr.ProdWidth
                        })
                    .Distinct()
                    .ToList();
                dbContext.Dispose();
                return prodQualityReports;
            }
        }
    }
}
