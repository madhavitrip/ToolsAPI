using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace ToolsAPI.Models
{
    // ==========================================
    // 1. DATABASE ENTITY
    // ==========================================
    [Table("applicationmessages")]
    public class ApplicationMessage
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public int Id { get; set; }

        [Required]
        [MaxLength(150)]
        public string MessageKey { get; set; } = string.Empty;

        [Required]
        [MaxLength(50)]
        public string Module { get; set; } = string.Empty;

        [MaxLength(255)]
        public string? Description { get; set; }

        [Required]
        public string EnglishText { get; set; } = string.Empty;

        public string? HindiText { get; set; }

        public bool IsActive { get; set; } = true;

        public DateTime CreatedAt { get; set; } = DateTime.UtcNow;

        public DateTime UpdatedAt { get; set; } = DateTime.UtcNow;
    }

    // ==========================================
    // 2. DATA TRANSFER OBJECTS (DTOs)
    // ==========================================
    public class MessageDto
    {
        public int Id { get; set; }
        public string MessageKey { get; set; } = string.Empty;
        public string Module { get; set; } = string.Empty;
        public string? Description { get; set; }
        public string EnglishText { get; set; } = string.Empty;
        public string? HindiText { get; set; }
        public bool IsActive { get; set; }
        public DateTime CreatedAt { get; set; }
        public DateTime UpdatedAt { get; set; }
    }

    public class CreateMessageRequest
    {
        public string MessageKey { get; set; } = string.Empty;
        public string Module { get; set; } = string.Empty;
        public string? Description { get; set; }
        public string EnglishText { get; set; } = string.Empty;
        public string? HindiText { get; set; }
    }

    public class UpdateMessageRequest
    {
        public string? Description { get; set; }
        public string EnglishText { get; set; } = string.Empty;
        public string? HindiText { get; set; }
        public bool IsActive { get; set; }
    }

    public class UpdateStatusRequest
    {
        public bool IsActive { get; set; }
    }

    public class PagedResult<T>
    {
        public List<T> Items { get; set; } = new();
        public int TotalCount { get; set; }
        public int PageNumber { get; set; }
        public int PageSize { get; set; }
    }
}
