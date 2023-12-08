using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.EntityFrameworkCore;
using Npgsql;

namespace ShiftReportApp1
{
    [Table("product_cat")]
    public class ProductCat
    {
        [Key]
        [Column("product_id")]
        public int ProductID { get; set; }
        [Column("product_name")]
        public string ProductName { get; set; }
        [Column("density_min")]
        public float DensityMin { get; set; }
        [Column("density_max")]
        public float DensityMax { get; set; }
    }

    [Table("places_in_line")]
    public class PlaceInLine
    {
        [Key]
        [Column("places_id")]
        public int PlacesID { get; set; }
        [Column("places_name")]
        public string PlacesName { get; set; }
        [Column("section")]
        public string Section { get; set; }
    }

    [Table("shift_report")]
    public class ShiftReport
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [Column("shiftreport_id")]
        public int ShiftReportID { get; set; }
        [Column("shift_day")]
        public int ShiftDay { get; set; }
        [Column("shift_num")]
        public int ShiftNum { get; set; }
        [Column("shift_date")]
        public DateTime ShiftDate { get; set; }
    }

    [Table("defect_types")]
    public class DefectType
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [Column("defect_types_id")]
        public int DefectTypeID { get; set; }
        [Column("defect_name")]
        public string DefectName { get; set; }
    }

    [Table("stop_types")]
    public class StopType
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [Column("stop_types_id")]
        public int StopTypeID { get; set; }
        [Column("stop_name")]
        public string StopName { get; set; }
        [Column("stop_category")]
        public string StopCategory { get; set; }
    }

    [Table("prod_quality_report")]
    public class ProdQualityReport
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [Column("pqreport_id")]
        public int PQReportID { get; set; }
        [ForeignKey("ShiftReportID")]
        [Column("report")]
        public int Report { get; set; }
        [ForeignKey("ProductID")]
        [Column("product")]
        public int Product { get; set; }
        [Column("unspecified")]
        public bool Unspecified { get; set; }
        [Column("regrading")]
        public bool Regarding { get; set; }
        [Column("prod_depth")]
        public int ProdDepth { get; set; }
        [Column("length")]
        public int ProdLength { get; set; }
        [Column("width")]
        public int ProdWidth { get; set; }
        [Column("volume_pack")]
        public float VolumePack { get; set; }
        [Column("avg_density")]
        public float AvgDensity { get; set; }
        [Column("pack_count")]
        public int PackCount { get; set; }
        [Column("date_rec_pqr")]
        public DateTime DateRecordPQR { get; set; }
        [Column("weight")]
        public float Weight { get; set; }
        [Column("volume_prod")]
        public float VolumeProduct { get; set; }
    }

    [Table("prod_defect_report")]
    public class ProdDefectReport
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [Column("defectreport_id")]
        public int DefectReportID { get; set; }
        [ForeignKey("PQReportID")]
        [Column("product_report")]
        public int ProductReport { get; set; }
        [ForeignKey("DefectTypeID")]
        [Column("defect_type")]
        public int DefectType { get; set; }
        [Column("defect_volume_pack")]
        public float DefectVolumePack { get; set; }
        [Column("defect_density")]
        public float DefectDensity { get; set; }
        [Column("defect_pack_count")]
        public int DefectPackCount { get; set; }
        [Column("defect_volume")]
        public float DefectVolume { get; set; }
        [Column("defect_weight")]
        public float DefectWeight { get; set; }
    }

    [Table("stops_report")]
    public class StopsReport
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [Column("stopreport_id")]
        public int StopReportID { get; set; }
        [ForeignKey("ShiftReportID")]
        [Column("shift_report")]
        public int ShiftReport { get; set; }
        [ForeignKey("StopTypeID")]
        [Column("stop_type")]
        public int StopType { get; set; }
        [Column("stop_first_time")]
        public string StopFirstTime { get; set; }
        [Column("stop_end_time")]
        public string StopEndTime { get; set; }
        [Column("comment_stop")]
        public string CommentStop { get; set; }
        [Column("duration_stop_min")]
        public int DurationStopMin { get; set; }
        [ForeignKey("PlacesID")]
        [Column("place_stop")]
        public int PlaceStop { get; set; }
        [Column("date_rec_sr")]
        public DateTime DateRecordSR { get; set; }
        [Column("breakdown_without_stop")]
        public bool BreakdownWithoutStop { get; set; }
        [Column("centrifuge")]
        public int Centrifuge { get; set; }
    }
    //-------------------------------------------------------------------------------------------------------------------------
    // Создание контекста
    public class ShiftReportDbContext : DbContext
    {
        public DbSet<ProductCat> ProductCategories { get; set; }
        public DbSet<PlaceInLine> PlaceInLine { get; set; }
        public DbSet<ShiftReport> ShiftReport { get; set; }
        public DbSet<DefectType> DefectTypes { get; set; }
        public DbSet<StopType> StopType { get; set; }
        public DbSet<ProdQualityReport> ProdQualityReport { get; set; }
        public DbSet<ProdDefectReport> ProdDefectReport { get; set; }
        public DbSet<StopsReport> StopsReport { get; set; }
        protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
        {
            DataBaseConnection dbConnection = new DataBaseConnection();
            NpgsqlConnection connection = dbConnection.GetConnection(); // Здесь вызываем ваш метод
            optionsBuilder.UseNpgsql(connection);
        }
    }
}
