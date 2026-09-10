using System.ComponentModel.DataAnnotations.Schema;
using System.ComponentModel.DataAnnotations;

namespace Tools.Models
{
    public class TemporaryNrDatas
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public int Id { get; set; }
        public int ProjectId { get; set; }
        public string? CourseName { get; set; }
        public string? SubjectName { get; set; }
        public string? CenterCode { get; set; }
        
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
        public bool Status { get; set; } = true;
        public string? District { get; set; }
        public int DistrictSort { get; set; }
    }
}
