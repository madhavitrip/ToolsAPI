using System.ComponentModel.DataAnnotations.Schema;
using System.ComponentModel.DataAnnotations;

namespace Tools.Models
{
    public class ExtrasConfiguration
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public int Id { get; set; }
        public int ProjectId { get; set; }
        public int ExtraType { get; set; }
        public string Mode { get; set; }
        public string Value { get; set; }
        public string EnvelopeType { get; set; }

    }
}
