using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace Tools.Models
{
    public class NewEnvelopeBreakingResult
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public int Id { get; set; }
        public int ProjectId { get; set; }
        public int NrDataId { get; set; }      // Always set - for extras, use the base NRData's NrDataId
       public int CenterListId { get; set; }
        public string EnvQuantity { get; set; }
        public int CenterEnv { get; set; }
        public int TotalEnv { get; set; }
        public string? Env { get; set; }         // "1/2", "2/2"
        public int SerialNumber { get; set; }
        public string? BookletSerial { get; set; }
        public string? OmrSerial { get; set; }
        public string? PackingDenomination { get; set; }
        public DateTime CreatedAt { get; set; } = DateTime.UtcNow;
        public bool Status { get; set; } = true;
    }
}
