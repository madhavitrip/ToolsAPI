using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace Tools.Models
{
    public class RPTTemplate
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public int TemplateId { get; set; }
        [Required]
        public int GroupId { get; set; }
        [Required]
        public int TypeId { get; set; }
        [Required]
        public string? TemplateName { get; set; }
        public string? RPTFilePath { get; set; }
        public string? ParsedFieldsJson { get; set; }
        public int Version { get; set; }
        public DateTime CreatedDate { get; set; } = DateTime.Now;
        public DateTime UpdatedDate { get; set;} = DateTime.Now;
        public bool IsActive { get; set; }
    }
}
