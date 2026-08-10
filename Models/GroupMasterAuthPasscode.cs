using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace ToolsAPI.Models
{
    [Table("GroupMasterAuthPasscodes")]
    public class GroupMasterAuthPasscode
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public int Id { get; set; }

        [Required]
        public int GroupId { get; set; }

        [Required]
        [MaxLength(256)]
        public string PasscodeHash { get; set; } = string.Empty;

        public int UpdatedByUserId { get; set; }

        public DateTime UpdatedAt { get; set; } = DateTime.UtcNow;
    }
}
