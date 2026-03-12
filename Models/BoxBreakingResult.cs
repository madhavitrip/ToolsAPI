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
        public int? NrDataId { get; set; }      // NULL if this is an extra
        public int? ExtraId { get; set; }       // NULL if regular NRData (1=Nodal, 2=University, 3=Office)
        public string CatchNo { get; set; }     // Store for reference

        // Box breaking results (calculated fields)
        public int Start { get; set; }
        public int End { get; set; }
        public string Serial { get; set; }
        public int TotalPages { get; set; }
        public string BoxNo { get; set; }
        public string OmrSerial { get; set; }
        public int? InnerBundlingSerial { get; set; }
        public int SerialNumber { get; set; }   // Serial number for this row

        // Metadata
        public DateTime CreatedAt { get; set; } = DateTime.UtcNow;
        public int UploadBatch { get; set; }
    }
}
