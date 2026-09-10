using System.ComponentModel.DataAnnotations.Schema;
using System.ComponentModel.DataAnnotations;

namespace Tools.Models
{
    public class NodalList
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public int Id { get; set; }
        public int ProjectId { get; set; }
        public int CollegeCode { get; set; }
        public string? CollegeName { get; set; }
        public int ExamCenterCode { get; set; }
        public string? ExamCenterName { get; set; }
        public string? Gender { get; set; }
        public int NodalCode { get; set; }
        public string? NodalName { get; set; }
        public string? OtherFields { get; set; }
    }
}
