using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace ERPToolsAPI.Models
{
    public class ExcelUpload
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public int Id { get; set; }
        public int GroupId { get; set; } // Foreign key to Group table
            public string? Type { get; set; }
            public string? Language { get; set; }
            public string? Subject { get; set; }
            public string? Course { get; set; }
            public string? ExamType { get; set; }
            public string? Catch { get; set; }
        public string? PaperNumber { get; set; }
        

    }
}
