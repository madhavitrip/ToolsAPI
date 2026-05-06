using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using Tools.Models;

namespace ToolsAPI.Models
{
    [Table("EnvelopeLotReports")]
    public class EnvelopeLotReport
    {
        [Key]
        public int Id { get; set; }

        [Required]
        public int ProjectId { get; set; }

        [Required]
        public int TemplateId { get; set; }

        [Required]
        public string TemplateName { get; set; } = string.Empty;

        [Required]
        public string EnvLotNumbers { get; set; } = string.Empty; // Comma-separated envelope lot numbers

        [Required]
        public string FileName { get; set; } = string.Empty;

        [Required]
        public DateTime GeneratedAt { get; set; }

        [Required]
        public string GeneratedBy { get; set; } = string.Empty;

        public string? FilePath { get; set; } // Make nullable to handle NULL values from database

        // Navigation properties
        [ForeignKey("ProjectId")]
        public virtual Project? Project { get; set; }
    }
}