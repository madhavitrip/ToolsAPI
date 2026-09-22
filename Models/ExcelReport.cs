using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace Tools.Models
{
    public class ExcelReport
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public int Id { get; set; }
        public int? ProjectId { get; set; }
        public int ModuleId { get; set; }
        public int Version { get; set; }
        public int? Lot { get; set; }
        public DateTime GeneratedAt { get; set; } = DateTime.Now;
        public int? GeneratedByUserId { get; set; }
        public string? FilePath { get; set; }
        public bool Status { get; set; }
    }
}
