using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace Tools.Models
{
    public class BoxBreakingResult
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public int Id { get; set; }

        public int ProjectId { get; set; }
        public int? EnvelopeBreakingResultId { get; set; }  // Reference to EnvelopeBreakingResults (nullable for split rows)

        // Box breaking results (calculated fields)
        public int Start { get; set; }
        public int End { get; set; }
        public string Serial { get; set; }
        public int TotalPages { get; set; }
        public string BoxNo { get; set; }
        public string OmrSerial { get; set; }
        public string BookletSerial { get; set; }
        public int? InnerBundlingSerial { get; set; }
        public int Quantity { get; set; }       // Quantity for this box
        public int Status { get; set; }
        // Metadata
        public DateTime CreatedAt { get; set; } = DateTime.UtcNow;
        public int UploadBatch { get; set; }
    }
}
