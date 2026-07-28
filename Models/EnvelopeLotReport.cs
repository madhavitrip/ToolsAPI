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

        // Removed [Required] to allow for project-wide reports or empty lot selections
        public string? EnvLotNumbers { get; set; } // Comma-separated envelope lot numbers (null or "0" for non-envelope-dependent templates)

        public int? LotNo { get; set; } // The actual lot number for lot-based reports (e.g., box breaking)

        [Required]
        public string FileName { get; set; } = string.Empty;

        [Required]
        public DateTime GeneratedAt { get; set; }

        public int? GeneratedByUserId { get; set; } // Store the user ID who generated the report

        public string? FilePath { get; set; } // Make nullable to handle NULL values from database

        public int? DownloadedByUserId { get; set; } // Store the user ID who downloaded the report

        public DateTime? DownloadedAt { get; set; } // Track when the report was last downloaded

        // Navigation properties
        [ForeignKey("ProjectId")]
        public virtual Project? Project { get; set; }

    }
}