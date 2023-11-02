using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Cache;
using System.Text;
using System.Threading.Tasks;

namespace ShiftReportApp1
{
    public class Names
    {
        public static string Request(int getstring)
        {
            if (getstring == 1)
            {
                return @"
                SELECT DISTINCT
                    pc.product_name,
                    pqr.unspecified,
                    pqr.prod_depth,
                    pqr.length,
                    pqr.width
                FROM prod_quality_report pqr
                JOIN product_cat pc ON pqr.product = pc.product_id
                JOIN shift_report sr ON pqr.report = sr.shiftreport_id
                WHERE sr.shift_date = @varDate1 AND sr.shift_day = 2;
                "; // утром за предыдущую смену
            }
            //-------------------------------------------------------------------------------------------------------------------
            else if (getstring == 2)
            {
                return @"
                SELECT DISTINCT 
                    pc.product_name,
                    pqr.unspecified,
                    pqr.prod_depth,
                    pqr.length,
                    pqr.width
                FROM prod_quality_report pqr
                JOIN product_cat pc ON pqr.product = pc.product_id
                JOIN shift_report sr ON pqr.report = sr.shiftreport_id
                WHERE sr.shift_date = @varDate1 AND sr.shift_day = 1;
                "; // вечером за предыдущую смену
            }
            //-------------------------------------------------------------------------------------------------------------------
            else if (getstring == 3)
            {
                return @"
                SELECT DISTINCT
                    pc.product_name,
                    pqr.unspecified,
                    pqr.prod_depth,
                    pqr.length,
                    pqr.width
                FROM prod_quality_report pqr
                JOIN product_cat pc ON pqr.product = pc.product_id
                JOIN shift_report sr ON pqr.report = sr.shiftreport_id
                WHERE sr.shift_date = @varDate1 AND sr.shift_day = 1;
                "; // утром до 8-00 за предыдущую смену
            }
            //-------------------------------------------------------------------------------------------------------------------
            else if (getstring == 4)
            {
                return @"
                SELECT DISTINCT pc.product_name, pqr.unspecified, pqr.prod_depth, pqr.length, pqr.width
                FROM prod_quality_report pqr
                JOIN product_cat pc ON pqr.product = pc.product_id
                JOIN shift_report sr ON pqr.report = sr.shiftreport_id
                WHERE sr.shift_date = @varDate1
                "; // продукты за предыдущие сутки
            }
            //-------------------------------------------------------------------------------------------------------------------
            else if (getstring == 5)
            {
                return @"
                SELECT DISTINCT pc.product_name, pqr.unspecified, pqr.prod_depth, pqr.length, pqr.width
                FROM prod_quality_report pqr
                JOIN product_cat pc ON pqr.product = pc.product_id
                JOIN shift_report sr ON pqr.report = sr.shiftreport_id
                WHERE sr.shift_date BETWEEN @varDate1 AND @varDate2;
                "; // продукты за период с даты по дату
            }
            //-------------------------------------------------------------------------------------------------------------------
            else if (getstring == 6)
            {
                return @"
                SELECT DISTINCT
                    pc.product_name, 
                    pqr.unspecified, 
                    pqr.prod_depth, 
                    pqr.length, 
                    pqr.width,
                    AVG(pqr.avg_dencity) AS avg_density_avg,
                    SUM(pqr.pack_count) AS pack_count_sum,
                    SUM(pqr.volume_prod) AS volume_prod_sum,
                    SUM(pqr.weight) AS weight_sum,
                    SUM(CASE WHEN dr.defect_type BETWEEN 1 AND 7 THEN dr.defect_pack_count ELSE 0 END) as low_qual_count,
                    SUM(CASE WHEN dr.defect_type BETWEEN 1 AND 7 THEN dr.defect_volume ELSE 0 END) as low_qual_vol,
                    SUM(CASE WHEN dr.defect_type BETWEEN 1 AND 7 THEN dr.defect_weight ELSE 0 END) as low_qual_weight,
                    SUM(CASE WHEN dr.defect_type = 8 THEN dr.defect_volume ELSE 0 END) as reject_vol,
                    SUM(CASE WHEN dr.defect_type = 8 THEN dr.defect_weight ELSE 0 END) as reject_weight,
                    (SUM(CASE WHEN dr.defect_type BETWEEN 1 AND 7 THEN dr.defect_weight ELSE 0 END) / SUM(pqr.weight)) * 100 as percent_low_qual,
                    (SUM(CASE WHEN dr.defect_type = 8 THEN dr.defect_weight ELSE 0 END) / SUM(pqr.weight)) * 100 as percent_reject
                FROM prod_quality_report pqr
                JOIN product_cat pc ON pqr.product = pc.product_id
                JOIN shift_report sr ON pqr.report = sr.shiftreport_id
                JOIN prod_defect_report dr ON pqr.pqreport_id = dr.product_report
                WHERE sr.shift_date BETWEEN @varDate1 AND @varDate2
                GROUP BY 
                    pc.product_name, 
                    pqr.unspecified, 
                    pqr.prod_depth, 
                    pqr.length, 
                    pqr.width;
                ";// отчет по качеству за период с даты по дату

            }
            //-------------------------------------------------------------------------------------------------------------------
            else if (getstring == 7)
            {
                return @"
                SELECT DISTINCT
                    sr.shift_date,
                    sr.shift_num,
                    pqr.volume_pack,
                    pc.product_name,
                    pqr.unspecified,
                    pqr.prod_depth,
                    pqr.length,
                    pqr.width,
                    AVG(pqr.avg_dencity) AS avg_density_avg,
                    SUM(pqr.pack_count) AS pack_count_sum,
                    SUM(pqr.volume_prod) AS volume_prod_sum,
                    SUM(pqr.weight) AS weight_sum,
                    SUM(CASE WHEN dr.defect_type BETWEEN 1 AND 7 THEN dr.defect_pack_count ELSE 0 END) as low_qual_count,
                    SUM(CASE WHEN dr.defect_type BETWEEN 1 AND 7 THEN dr.defect_volume ELSE 0 END) as low_qual_vol,
                    SUM(CASE WHEN dr.defect_type BETWEEN 1 AND 7 THEN dr.defect_weight ELSE 0 END) as low_qual_weight,
                    SUM(CASE WHEN dr.defect_type = 8 THEN dr.defect_volume ELSE 0 END) as reject_vol,
                    SUM(CASE WHEN dr.defect_type = 8 THEN dr.defect_weight ELSE 0 END) as reject_weight,
                    (SUM(CASE WHEN dr.defect_type BETWEEN 1 AND 7 THEN dr.defect_weight ELSE 0 END) / SUM(pqr.weight)) * 100 as percent_low_qual,
                    (SUM(CASE WHEN dr.defect_type = 8 THEN dr.defect_weight ELSE 0 END) / SUM(pqr.weight)) * 100 as percent_reject
                FROM prod_quality_report pqr
                JOIN product_cat pc ON pqr.product = pc.product_id
                JOIN shift_report sr ON pqr.report = sr.shiftreport_id
                JOIN prod_defect_report dr ON pqr.pqreport_id = dr.product_report
                WHERE (sr.shift_date = @varDate1) AND sr.shift_num IN(@numShift1, @numShift2, @numShift3, @numShift4)
                GROUP BY 
                    sr.shift_date,
                    sr.shift_num,
                    pc.product_name, 
                    pqr.unspecified, 
                    pqr.prod_depth, 
                    pqr.length, 
                    pqr.width,
					pqr.volume_pack,
					pc.product_name,
                    pqr.unspecified
                ORDER BY
                    sr.shift_num;

                "; // отчет за предыдущие сутки с разбивкой по сменам
            }
            //-------------------------------------------------------------------------------------------------------------------
            else if (getstring == 8)
            {
                return @"
                SELECT
                    sr.shift_date,
                    sr.shift_num,
                    pqr.volume_pack,
                    pc.product_name, 
                    pqr.unspecified, 
                    pqr.prod_depth, 
                    pqr.length, 
                    pqr.width,
                    AVG(pqr.avg_dencity) AS avg_density_avg,
                    SUM(pqr.pack_count) AS pack_count_sum,
                    SUM(pqr.volume_prod) AS volume_prod_sum,
                    SUM(pqr.weight) AS weight_sum,
                    SUM(CASE WHEN dr.defect_type BETWEEN 1 AND 7 THEN dr.defect_pack_count ELSE 0 END) as low_qual_count,
                    SUM(CASE WHEN dr.defect_type BETWEEN 1 AND 7 THEN dr.defect_volume ELSE 0 END) as low_qual_vol,
                    SUM(CASE WHEN dr.defect_type BETWEEN 1 AND 7 THEN dr.defect_weight ELSE 0 END) as low_qual_weight,
                    SUM(CASE WHEN dr.defect_type = 8 THEN dr.defect_volume ELSE 0 END) as reject_vol,
                    SUM(CASE WHEN dr.defect_type = 8 THEN dr.defect_weight ELSE 0 END) as reject_weight,
                    (SUM(CASE WHEN dr.defect_type BETWEEN 1 AND 7 THEN dr.defect_weight ELSE 0 END) / SUM(pqr.weight)) * 100 as percent_low_qual,
                    (SUM(CASE WHEN dr.defect_type = 8 THEN dr.defect_weight ELSE 0 END) / SUM(pqr.weight)) * 100 as percent_reject
                FROM prod_quality_report pqr
                JOIN product_cat pc ON pqr.product = pc.product_id
                JOIN shift_report sr ON pqr.report = sr.shiftreport_id
                JOIN prod_defect_report dr ON pqr.pqreport_id = dr.product_report
                WHERE (sr.shift_date = @varDate1) AND sr.shift_day = 2
                GROUP BY
                    sr.shift_date,
                    sr.shift_num,
                    pqr.volume_pack,
                    pc.product_name, 
                    pqr.unspecified, 
                    pqr.prod_depth, 
                    pqr.length, 
                    pqr.width;
                "; // отчет за предыдущую смену с 9-00 до 21-00
            }
            //-------------------------------------------------------------------------------------------------------------------
            else if (getstring == 9)
            {
                return @"
                SELECT
                    sr.shift_date,
                    sr.shift_num,
                    pqr.volume_pack,
                    pc.product_name, 
                    pqr.unspecified, 
                    pqr.prod_depth, 
                    pqr.length, 
                    pqr.width,
                    AVG(pqr.avg_dencity) AS avg_density_avg,
                    SUM(pqr.pack_count) AS pack_count_sum,
                    SUM(pqr.volume_prod) AS volume_prod_sum,
                    SUM(pqr.weight) AS weight_sum,
                    SUM(CASE WHEN dr.defect_type BETWEEN 1 AND 7 THEN dr.defect_pack_count ELSE 0 END) as low_qual_count,
                    SUM(CASE WHEN dr.defect_type BETWEEN 1 AND 7 THEN dr.defect_volume ELSE 0 END) as low_qual_vol,
                    SUM(CASE WHEN dr.defect_type BETWEEN 1 AND 7 THEN dr.defect_weight ELSE 0 END) as low_qual_weight,
                    SUM(CASE WHEN dr.defect_type = 8 THEN dr.defect_volume ELSE 0 END) as reject_vol,
                    SUM(CASE WHEN dr.defect_type = 8 THEN dr.defect_weight ELSE 0 END) as reject_weight,
                    (SUM(CASE WHEN dr.defect_type BETWEEN 1 AND 7 THEN dr.defect_weight ELSE 0 END) / SUM(pqr.weight)) * 100 as percent_low_qual,
                    (SUM(CASE WHEN dr.defect_type = 8 THEN dr.defect_weight ELSE 0 END) / SUM(pqr.weight)) * 100 as percent_reject
                FROM prod_quality_report pqr
                JOIN product_cat pc ON pqr.product = pc.product_id
                JOIN shift_report sr ON pqr.report = sr.shiftreport_id
                JOIN prod_defect_report dr ON pqr.pqreport_id = dr.product_report
                WHERE (sr.shift_date = @varDate1) AND sr.shift_day = 1
                GROUP BY
                    sr.shift_date,
                    sr.shift_num,
                    pqr.volume_pack,
                    pc.product_name, 
                    pqr.unspecified, 
                    pqr.prod_depth, 
                    pqr.length, 
                    pqr.width;
                "; // отчет за предыдущую смену запрос с 00-00 до 9-00
            }
            //-------------------------------------------------------------------------------------------------------------------
            else if (getstring == 10)
            {
                return @"
                SELECT
                    sr.shift_date,
                    sr.shift_num,
                    pqr.volume_pack,
                    pc.product_name, 
                    pqr.unspecified, 
                    pqr.prod_depth, 
                    pqr.length, 
                    pqr.width,
                    AVG(pqr.avg_dencity) AS avg_density_avg,
                    SUM(pqr.pack_count) AS pack_count_sum,
                    SUM(pqr.volume_prod) AS volume_prod_sum,
                    SUM(pqr.weight) AS weight_sum,
                    SUM(CASE WHEN dr.defect_type BETWEEN 1 AND 7 THEN dr.defect_pack_count ELSE 0 END) as low_qual_count,
                    SUM(CASE WHEN dr.defect_type BETWEEN 1 AND 7 THEN dr.defect_volume ELSE 0 END) as low_qual_vol,
                    SUM(CASE WHEN dr.defect_type BETWEEN 1 AND 7 THEN dr.defect_weight ELSE 0 END) as low_qual_weight,
                    SUM(CASE WHEN dr.defect_type = 8 THEN dr.defect_volume ELSE 0 END) as reject_vol,
                    SUM(CASE WHEN dr.defect_type = 8 THEN dr.defect_weight ELSE 0 END) as reject_weight,
                    (SUM(CASE WHEN dr.defect_type BETWEEN 1 AND 7 THEN dr.defect_weight ELSE 0 END) / SUM(pqr.weight)) * 100 as percent_low_qual,
                    (SUM(CASE WHEN dr.defect_type = 8 THEN dr.defect_weight ELSE 0 END) / SUM(pqr.weight)) * 100 as percent_reject
                FROM prod_quality_report pqr
                JOIN product_cat pc ON pqr.product = pc.product_id
                JOIN shift_report sr ON pqr.report = sr.shiftreport_id
                JOIN prod_defect_report dr ON pqr.pqreport_id = dr.product_report
                WHERE (sr.shift_date = @varDate1) AND sr.shift_day = 1
                GROUP BY
                    sr.shift_date,
                    sr.shift_num,
                    pqr.volume_pack,
                    pc.product_name, 
                    pqr.unspecified, 
                    pqr.prod_depth, 
                    pqr.length, 
                    pqr.width;
                "; // отчет за предыдущую смену с 21-00 до 23-59
            }
            //-------------------------------------------------------------------------------------------------------------------
            else if (getstring == 11)
            {
                return @"
                SELECT DISTINCT
                    sr.shift_num,
                    pc.product_name,
                    pqr.unspecified,
                    pqr.prod_depth,
                    pqr.length,
                    pqr.width,
                    AVG(pqr.avg_dencity) AS avg_density_avg,
                    SUM(pqr.pack_count) AS pack_count_sum,
                    SUM(pqr.volume_prod) AS volume_prod_sum,
                    SUM(pqr.weight) AS weight_sum,
                    SUM(CASE WHEN dr.defect_type BETWEEN 1 AND 7 THEN dr.defect_pack_count ELSE 0 END) as low_qual_count,
                    SUM(CASE WHEN dr.defect_type BETWEEN 1 AND 7 THEN dr.defect_volume ELSE 0 END) as low_qual_vol,
                    SUM(CASE WHEN dr.defect_type BETWEEN 1 AND 7 THEN dr.defect_weight ELSE 0 END) as low_qual_weight,
                    SUM(CASE WHEN dr.defect_type = 8 THEN dr.defect_volume ELSE 0 END) as reject_vol,
                    SUM(CASE WHEN dr.defect_type = 8 THEN dr.defect_weight ELSE 0 END) as reject_weight
                FROM prod_quality_report pqr
                JOIN product_cat pc ON pqr.product = pc.product_id
                JOIN shift_report sr ON pqr.report = sr.shiftreport_id
                JOIN prod_defect_report dr ON pqr.pqreport_id = dr.product_report
                WHERE (sr.shift_date BETWEEN @varDate1 AND @varDate2) AND sr.shift_num IN(@numShift1, @numShift2, @numShift3, @numShift4)
                GROUP BY 
                    sr.shift_num,
                    pc.product_name, 
                    pqr.unspecified, 
                    pqr.prod_depth, 
                    pqr.length, 
                    pqr.width
                ORDER BY
                    sr.shift_num;
                "; // отчет с даты по дату с разбивкой по номерам смен 
            }
            //-------------------------------------------------------------------------------------------------------------------
            else if (getstring == 12)
            {
                return @"
                SELECT DISTINCT
                    dr.defect_type AS ID,
                    dt.type_name AS Дефекты,
                    SUM(dr.defect_pack_count) AS Количество_упаковок,
                    SUM(dr.defect_volume) AS Объем,
                    SUM(dr.defect_weight) AS Вес,
                    (SUM(CASE WHEN dr.defect_type BETWEEN 1 AND 7 THEN dr.defect_weight ELSE 0 END) /
                        (SELECT SUM(CASE WHEN dr.defect_type BETWEEN 1 AND 7 THEN dr.defect_weight ELSE 0 END)
                        FROM prod_quality_report pqr
                        JOIN shift_report sr ON pqr.report = sr.shiftreport_id
                        JOIN prod_defect_report dr ON pqr.pqreport_id = dr.product_report
                        WHERE(sr.shift_date BETWEEN @varDate1 AND @varDate2) AND sr.shift_num IN(@numShift1, @numShift2, @numShift3, @numShift4)
                        GROUP BY
                            sr.shift_num)) * 100 AS Процент_вида_ОС,
                    (SUM(CASE WHEN dr.defect_type = 8 THEN dr.defect_weight ELSE 0 END) /
                        (SELECT SUM(CASE WHEN dr.defect_type = 8 THEN dr.defect_weight ELSE 0 END)
                        FROM prod_quality_report pqr
                        JOIN shift_report sr ON pqr.report = sr.shiftreport_id
                        JOIN prod_defect_report dr ON pqr.pqreport_id = dr.product_report
                        WHERE(sr.shift_date BETWEEN @varDate1 AND @varDate2) AND sr.shift_num IN(@numShift1, @numShift2, @numShift3, @numShift4)
                        GROUP BY
                            sr.shift_num)) * 100 AS Процент_брака
                FROM prod_quality_report pqr
                JOIN shift_report sr ON pqr.report = sr.shiftreport_id
                JOIN prod_defect_report dr ON pqr.pqreport_id = dr.product_report
                JOIN defect_types dt ON dr.defect_type = dt.defect_types_id
                WHERE sr.shift_date BETWEEN @varDate1 AND @varDate2
                GROUP BY 
                    dr.defect_type,
                    dt.type_name
                ORDER BY
                    dr.defect_type;
                "; // отчет с даты по дату по типам брака
            }
            //-------------------------------------------------------------------------------------------------------------------
            else if (getstring == 13)
            {
                return @"
                SELECT DISTINCT
                    sr.shift_num AS Смена,
                    dr.defect_type AS ID,
                    dt.type_name AS Дефекты,
                    SUM(dr.defect_pack_count) AS Количество_упаковок,
                    SUM(dr.defect_volume) AS Объем,
                    SUM(dr.defect_weight) AS Вес,
                    (SUM(CASE WHEN dr.defect_type BETWEEN 1 AND 7 THEN dr.defect_weight ELSE 0 END) /
                        (SELECT SUM(CASE WHEN dr.defect_type BETWEEN 1 AND 7 THEN dr.defect_weight ELSE 0 END)
                        FROM prod_quality_report pqr
                        JOIN shift_report sr ON pqr.report = sr.shiftreport_id
                        JOIN prod_defect_report dr ON pqr.pqreport_id = dr.product_report
                        WHERE(sr.shift_date BETWEEN @varDate1 AND @varDate2) AND sr.shift_num IN(@numShift1, @numShift2, @numShift3, @numShift4)
                        GROUP BY
                            sr.shift_num)) * 100 AS Процент_вида_ОС,
                    (SUM(CASE WHEN dr.defect_type = 8 THEN dr.defect_weight ELSE 0 END) /
                        (SELECT SUM(CASE WHEN dr.defect_type = 8 THEN dr.defect_weight ELSE 0 END)
                        FROM prod_quality_report pqr
                        JOIN shift_report sr ON pqr.report = sr.shiftreport_id
                        JOIN prod_defect_report dr ON pqr.pqreport_id = dr.product_report
                        WHERE(sr.shift_date BETWEEN @varDate1 AND @varDate2) AND sr.shift_num IN(@numShift1, @numShift2, @numShift3, @numShift4)
                        GROUP BY
                            sr.shift_num)) * 100 AS Процент_брака
                FROM prod_quality_report pqr
                JOIN shift_report sr ON pqr.report = sr.shiftreport_id
                JOIN prod_defect_report dr ON pqr.pqreport_id = dr.product_report
                JOIN defect_types dt ON dr.defect_type = dt.defect_types_id
                WHERE (sr.shift_date BETWEEN @varDate1 AND @varDate2) AND sr.shift_num IN(@numShift1, @numShift2, @numShift3, @numShift4)
                GROUP BY
                    sr.shift_num,
                    dr.defect_type,
                    dt.type_name
                ORDER BY
                    sr.shift_num,
                    dr.defect_type;
                "; // отчет с даты по дату по типам брака с разбивкой по номерам смен 
            }
            //-------------------------------------------------------------------------------------------------------------------
            else if (getstring == 21)
            {
                return @"
                SELECT
                    st.type_index AS Типы_простоев,
                    st.type_name AS Название_остановки, 
                    pil.places_name AS Место, 
                    SUM(str.duration_stop_min) AS Время_простоя
                FROM stops_report str
                JOIN stop_types st ON str.stop_type = st.stop_types_id
                JOIN shift_report sr ON str.report = sr.shiftreport_id
                JOIN places_in_line pil ON str.place_stop = pil.places_id
                WHERE sr.shift_date BETWEEN @varDate1 AND @varDate2
                GROUP BY st.type_index, st.type_name, pil.places_name
                ORDER BY SUM(str.duration_stop_min) DESC;
                "; // отчет с даты по дату по простоям
            }
            //-------------------------------------------------------------------------------------------------------------------
            else if (getstring == 22)
            {
                return @"
                SELECT 
                    st.type_index AS Типы_простоев, 
                    SUM(str.duration_stop_min) AS Время_простоя
                FROM stops_report str
                JOIN stop_types st ON str.stop_type = st.stop_types_id
                JOIN shift_report sr ON str.report = sr.shiftreport_id
                JOIN places_in_line pil ON str.place_stop = pil.places_id
                WHERE (sr.shift_date BETWEEN @varDate1 AND @varDate2) AND st.type_index IN(@typeStops1, @typeStops2, @typeStops3, @typeStops4)
                GROUP BY st.type_index
                ORDER BY SUM(str.duration_stop_min) DESC;
                "; // отчет с даты по дату по простоям с разбивкой по видам простоя
            }
            //-------------------------------------------------------------------------------------------------------------------
            else if (getstring == 23)
            {
                return @"
                SELECT
                    sr.shift_num AS Номер_смены,
                    st.type_index AS Типы_простоев, 
                    st.type_name AS Название_остановки,
                    pil.places_name AS Место, 
                    SUM(str.duration_stop_min) AS Время_простоя
                FROM stops_report str
                JOIN stop_types st ON str.stop_type = st.stop_types_id
                JOIN shift_report sr ON str.report = sr.shiftreport_id
                JOIN places_in_line pil ON str.place_stop = pil.places_id
                WHERE (sr.shift_date BETWEEN @varDate1 AND @varDate2) AND sr.shift_num IN(@numShift1, @numShift2, @numShift3, @numShift4)
                GROUP BY sr.shift_num, st.type_index, st.type_name, pil.places_name
                ORDER BY sr.shift_num, SUM(str.duration_stop_min) DESC;
                "; // отчет с даты по дату по простоям с разбивкой по номерам смен
            }
            //-------------------------------------------------------------------------------------------------------------------
            else if (getstring == 24)
            {
                return @"
                SELECT
                    sr.shift_num AS Номер_смены,
                    st.type_index AS Типы_простоев,
                    pil.places_name AS Место, 
                    SUM(str.duration_stop_min) AS Время_простоя
                FROM stops_report str
                JOIN stop_types st ON str.stop_type = st.stop_types_id
                JOIN shift_report sr ON str.report = sr.shiftreport_id
                JOIN places_in_line pil ON str.place_stop = pil.places_id
                WHERE (sr.shift_date BETWEEN @varDate1 AND @varDate2) AND sr.shift_num IN(@numShift1, @numShift2, @numShift3, @numShift4) AND st.type_index IN(@typeStops1, @typeStops2, @typeStops3, @typeStops4)
                GROUP BY  sr.shift_num, st.type_index, pil.places_name
                ORDER BY SUM(str.duration_stop_min) DESC;
                "; // отчет с даты по дату по простоям с разбивкой по видам простоя по номерам смен
            }
            //-------------------------------------------------------------------------------------------------------------------
            else if (getstring == 25)
            {
                return @"
                SELECT
                    sr.shift_date AS Дата,
                    st.type_index AS Типы_простоев,
                    st.type_name AS Название_остановки, 
                    pil.places_name AS Место, 
                    str.duration_stop_min AS Время_простоя,
                    str.comment_stop AS Комментарий
                FROM stops_report str
                JOIN stop_types st ON str.stop_type = st.stop_types_id
                JOIN shift_report sr ON str.report = sr.shiftreport_id
                JOIN places_in_line pil ON str.place_stop = pil.places_id
                WHERE sr.shift_date BETWEEN @varDate1 AND @varDate2
                GROUP BY sr.shift_date, st.type_index, st.type_name, pil.places_name, str.duration_stop_min ,str.comment_stop 
                ORDER BY sr.shift_date;

                "; // отчет с даты по дату с комментариями по простоям
            }
            //-------------------------------------------------------------------------------------------------------------------
            else if (getstring == 26)
            {
                return @"
                SELECT
                    sr.shift_num, 
                    st.type_index AS Типы_простоев,
                    st.type_name AS Название_остановки, 
                    pil.places_name AS Место, 
                    SUM(str.duration_stop_min) AS Время_простоя
                FROM stops_report str
                JOIN stop_types st ON str.stop_type = st.stop_types_id
                JOIN shift_report sr ON str.report = sr.shiftreport_id
                JOIN places_in_line pil ON str.place_stop = pil.places_id
                WHERE sr.shift_date = @varDate1 AND sr.shift_day = 1
                GROUP BY  st.type_index,st.type_name, pil.places_name
                ORDER BY SUM(str.duration_stop_min) DESC;
                "; // отчет с даты по дату по простоям
            }
            //-------------------------------------------------------------------------------------------------------------------
            else if (getstring == 27)
            {
                return @"
                SELECT
                    sr.shift_num, 
                    st.type_index AS Типы_простоев,
                    st.type_name AS Название_остановки, 
                    pil.places_name AS Место, 
                    SUM(str.duration_stop_min) AS Время_простоя
                FROM stops_report str
                JOIN stop_types st ON str.stop_type = st.stop_types_id
                JOIN shift_report sr ON str.report = sr.shiftreport_id
                JOIN places_in_line pil ON str.place_stop = pil.places_id
                WHERE sr.shift_date = @varDate1 AND sr.shift_day = 2
                GROUP BY  st.type_index, st.type_name, pil.places_name
                ORDER BY SUM(str.duration_stop_min) DESC;
                "; // отчет с даты по дату по простоям
            }
            //-------------------------------------------------------------------------------------------------------------------
            else if (getstring == 28)
            {
                return @"
                SELECT 
                    st.type_index AS Типы_простоев, 
                    pil.places_name AS Место, 
                    SUM(str.duration_stop_min) AS Время_простоя
                FROM stops_report str
                JOIN stop_types st ON str.stop_type = st.stop_types_id
                JOIN shift_report sr ON str.report = sr.shiftreport_id
                JOIN places_in_line pil ON str.place_stop = pil.places_id
                WHERE (sr.shift_date = @varDate1 AND sr.shift_day = 1) AND st.type_index IN(@typeStops1, @typeStops2, @typeStops3, @typeStops4)
                GROUP BY st.type_index, pil.places_name, st.type_index
                ORDER BY SUM(str.duration_stop_min) DESC;
                "; // отчет с даты по дату по простоям с разбивкой по видам простоя
            }
            //-------------------------------------------------------------------------------------------------------------------
            else if (getstring == 29)
            {
                return @"
                SELECT
                    st.type_index AS Типы_простоев, 
                    pil.places_name AS Место, 
                    SUM(str.duration_stop_min) AS Время_простоя
                FROM stops_report str
                JOIN stop_types st ON str.stop_type = st.stop_types_id
                JOIN shift_report sr ON str.report = sr.shiftreport_id
                JOIN places_in_line pil ON str.place_stop = pil.places_id
                WHERE (sr.shift_date = @varDate1 AND sr.shift_day = 2) AND st.type_index IN(@typeStops1, @typeStops2, @typeStops3, @typeStops4)
                GROUP BY st.type_index, pil.places_name, st.type_index
                ORDER BY SUM(str.duration_stop_min) DESC;
                "; // отчет с даты по дату по простоям с разбивкой по видам простоя
            }
            //-------------------------------------------------------------------------------------------------------------------
            else if (getstring == 30)
            {
                return @"
                SELECT
                    st.type_index AS Типы_простоев, 
                    pil.places_name AS Место, 
                    SUM(str.duration_stop_min) AS Время_простоя
                FROM stops_report str
                JOIN stop_types st ON str.stop_type = st.stop_types_id
                JOIN shift_report sr ON str.report = sr.shiftreport_id
                JOIN places_in_line pil ON str.place_stop = pil.places_id
                WHERE (sr.shift_date = @varDate1 AND sr.shift_day = 2) AND st.type_index IN(@typeStops1, @typeStops2, @typeStops3, @typeStops4)
                GROUP BY st.type_index, pil.places_name, st.type_index
                ORDER BY SUM(str.duration_stop_min) DESC;
                "; // отчет с даты по дату по простоям с разбивкой по видам простоя
            }
            //-------------------------------------------------------------------------------------------------------------------
            return null;
        }
    }
}
