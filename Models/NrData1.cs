using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace Tools.Models
{
    public class NrData1
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public int Id { get; set; }
        public int ProjectId { get; set; }
        public string? CatchNo { get; set; }
        public string? NRDatas { get; set; }
        public string? CourseName { get; set; }
        public string? SubjectName { get; set; }
        public string? ExamDate { get; set; }
        public string? ExamTime { get; set; }
        public string? Day { get; set; }
        public int Pages { get; set; }
        public int Steps { get; set; }
        public int EnvLotNo { get; set; }
        public int LotNo { get; set; }
        public int VerificationStatus { get; set; }
        public int Batch { get; set; }
        public string? Remarksss { get; set; }

    }
}
