using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace ToolsAPI.Models
{
    [Table("MasterAuthAuditLogs")]
    public class MasterAuthAuditLog
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public int Id { get; set; }

        public int UserId { get; set; }

        public int GroupId { get; set; }

        [Required]
        [MaxLength(100)]
        public string OperationType { get; set; } = string.Empty; // e.g. "CREATE", "UPDATE", "DELETE"

        [Required]
        [MaxLength(100)]
        public string ModuleName { get; set; } = string.Empty; // e.g. "RPTTemplates", "ProjectConfigs"

        [MaxLength(50)]
        public string IpAddress { get; set; } = string.Empty;

        public bool Success { get; set; }

        [MaxLength(500)]
        public string Message { get; set; } = string.Empty;

        public DateTime Timestamp { get; set; } = DateTime.UtcNow;
    }
}
