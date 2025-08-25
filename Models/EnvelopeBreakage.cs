using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace Tools.Models
{
    public class EnvelopeBreakage
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public int EnvelopeId { get; set; }
        public int ProjectId { get; set; }
        public int NrDataId { get; set; }
        public string InnerEnvelope {  get; set; }
        public string OuterEnvelope { get; set; }
    }
}
