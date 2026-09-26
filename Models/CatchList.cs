using System.ComponentModel.DataAnnotations.Schema;
using System.ComponentModel.DataAnnotations;

namespace Tools.Models
{
    public class CatchList
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public int Id { get; set; }
        public int ProjectId { get; set; }
        public string? CourseName { get; set; }
        public string? SubjectName { get; set; }
        public string? CollegeName { get; set; }
        [Required]
        public int CenterCode { get; set; }
        public string? CenterName { get; set; }
        public int CollegeCode { get; set; }
        public string? PaperCode { get; set; }
        [Required]
        public int NRQuantity { get; set; }
        [Required]
        public string? CatchNo { get; set; }
        public string? ExamDate { get; set; }
        public string? ExamTime { get; set; }
        public string? NRDatas { get; set; }
        public int Transgender { get; set; }
        public int Male { get; set; }
        public int Female { get; set; }
        public string? Semester { get; set; }
        public bool Status { get; set; } = true;
    }
}
