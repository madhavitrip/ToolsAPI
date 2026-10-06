using System.ComponentModel.DataAnnotations.Schema;
using System.ComponentModel.DataAnnotations;

namespace Tools.Models
{
    public class NewEnvelopeBreakage
    {
            [Key]
            [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
            public int EnvelopeId { get; set; }
            public int ProjectId { get; set; }
            public int NrDataId { get; set; }
            public int CenterListId { get; set; }
            public string InnerEnvelope { get; set; }
            public string OuterEnvelope { get; set; }
            public int TotalEnvelope { get; set; }
            public bool Status { get; set; } = true;
    }
}
