using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace Tools.Models
{
    public class RPTTemplate
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public int TemplateId { get; set; }
        public int? GroupId { get; set; }
        [Required]
        public int TypeId { get; set; }
        public int? ProjectId { get; set; }
        public int? UploadedByUserId { get; set; }
        public List<int>? ModuleIds { get; set; }
        [Required]
        public string? TemplateName { get; set; }
        public string? RPTFilePath { get; set; }
        public string? ParsedFieldsJson { get; set; }
        public string? DesignSnapshotJson { get; set; }
        public string? RequiredFieldsJson { get; set; }
        public int Version { get; set; }
        public DateTime CreatedDate { get; set; } = DateTime.Now;
        public DateTime UpdatedDate { get; set;} = DateTime.Now;
        public bool IsActive { get; set; }
        [NotMapped]
        public bool hasMapping { get; set; }
    }
}

