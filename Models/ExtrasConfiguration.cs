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
        public string Mode { get; set; } // Stores "Fixed" or "%"
        public string Value { get; set; }
        public string EnvelopeType { get; set; }
        public string? RangeConfig { get; set; } //Added by Akshaya to accept the range in json format
        public string? nodalValue { get; set; } // JSON format: [{"NodalCodes":"NC1,NC2,NC3","Value":"10"},{"NodalCodes":"NC4","Value":"20"}]

    }
}
