using System.ComponentModel.DataAnnotations.Schema;
using System.ComponentModel.DataAnnotations;

namespace Tools.Models
{
    public class NRData
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public int Id { get; set; }
        public int ProjectId { get; set; }
        public string? CourseName {  get; set; }
        public string? SubjectName { get; set; }
        public string? CenterCode { get; set; }
        public int? Quantity { get; set; }
        public string? CatchNo { get; set; }
        public string? ExamDate { get; set; }
        public string? ExamTime { get; set; }
        public string? NRDatas { get; set; }
        public string ? NodalCode { get; set; }
        public int Pages { get; set; }

    }
}
