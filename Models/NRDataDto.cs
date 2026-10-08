using System;
using System.Collections.Generic;

namespace ToolsAPI.Models
{
    public class NRDataDto
    {
        public int Id { get; set; }
        public int ProjectId { get; set; }
        public string? CourseName { get; set; }
        public string? SubjectName { get; set; }
        public string? CenterCode { get; set; }
        public int Quantity { get; set; }
        public int NRQuantity { get; set; }
        public string? CatchNo { get; set; }
        public string? ExamDate { get; set; }
        public string? ExamTime { get; set; }
        public string? Day { get; set; }
        public string? NRDatas { get; set; }
        public string? NodalCode { get; set; }
        public int Pages { get; set; }
        public string? Route { get; set; }
        public int RouteSort { get; set; }
        public double CenterSort { get; set; }
        public double NodalSort { get; set; }
        public string? Symbol { get; set; }
        public bool Status { get; set; }
        public int? NRDataId { get; set; }
        public int Steps { get; set; }
        public int LotNo { get; set; }
        public string? District { get; set; }
        public int DistrictSort { get; set; }
        public int EnvLotNo { get; set; }
        public int VerificationStatus { get; set; }
        public int Batch { get; set; }
        public string? Remarksss { get; set; }
    }
}
