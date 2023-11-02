using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ShiftReportApp1
{
    public class ProductCategory
    {
        public int ProductID { get; set; }
        public string ProductName { get; set; }
        public int DensityMin { get; set; }
        public int DensityMax { get; set; }
    }

    public class PlaceInLine
    {
        public int PlacesID { get; set; }
        public string PlacesName { get; set; }
        public string Section { get; set; }
    }

    public class ShiftReport
    {
        public int ShiftReportID { get; set; }
        public int ShiftDay { get; set; }
        public int ShiftNum { get; set; }
        public DateTime ShiftDate { get; set; }
    }

    public class DefectType
    {
        public int DefectTypeID { get; set; }
        public string TypeName { get; set; }
    }

    public class StopType
    {
        public int StopTypeID { get; set; }
        public string TypeName { get; set; }
        public string TypeIndex { get; set; }
    }

    public class ProdQualityReport
    {
        public int PQReportID { get; set; }
        public int Report { get; set; }
        public int Product { get; set; }
        public bool Unspecified { get; set; }
        public bool Regarding { get; set; }
        public int ProdDepth { get; set; }
        public int ProdLength { get; set; }
        public int ProdWidth { get; set; }
        public float VolumePack { get; set; }
        public float AvgDensity { get; set; }
        public int PackCount { get; set; }
        public DateTime DateRecordPQR { get; set; }
        public float Weight { get; set; }
        public float VolumeProduct { get; set; }
    }

    public class ProdDefectReport
    {
        public int DefectReportID { get; set; }
        public int ProductReport { get; set; }
        public int DefectType { get; set; }
        public int DefectVolumePack { get; set; }
        public float DefectDensity { get; set; }
        public int DefectPackCount { get; set; }
        public float DefectVolume { get; set; }
        public float DefectWeight { get; set; }
    }

    public class StopsReport
    {
        public int StopReportID { get; set; }
        public int Report { get; set; }
        public int StopType { get; set; }
        public TimeSpan StopFirstTime { get; set; }
        public TimeSpan StopEndTime { get; set; }
        public string CommentStop { get; set; }
        public int DurationStopMin { get; set; }
        public int PlaceStop { get; set; }
        public DateTime DateRecordSR { get; set; }
    }
}
