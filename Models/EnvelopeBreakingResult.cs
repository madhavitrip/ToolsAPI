using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace Tools.Models
{
    public class EnvelopeBreakingResult
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public int Id { get; set; }

        public int ProjectId { get; set; }
        public int? NrDataId { get; set; }      // NULL if this is an extra
        public int? ExtraId { get; set; }       // NULL if regular NRData (1=Nodal, 2=University, 3=Office)
        public string CatchNo { get; set; }     // Store for reference

        // Calculated fields from envelope breaking
        public int EnvQuantity { get; set; }
        public int CenterEnv { get; set; }
        public int TotalEnv { get; set; }
        public string Env { get; set; }         // "1/2", "2/2"
        public int SerialNumber { get; set; }   // Serial number (resets per CatchNo)
        public string BookletSerial { get; set; } // "1-50", "51-100"

        // ONLY for extras - modified sort values (NULL for regular NRData)
        public double? NodalSortModified { get; set; }   // For extras: modified NodalSort
        public int? CenterSortModified { get; set; }     // For extras: modified CenterSort
        public int? RouteSortModified { get; set; }      // For extras: modified RouteSort
        
        // ONLY for extras - reference to original NRData for metadata
        public string NodalCodeRef { get; set; }         // For extras: which NodalCode this extra belongs to
        public string RouteRef { get; set; }             // For extras: which Route this extra belongs to

        // Metadata
        public DateTime CreatedAt { get; set; } = DateTime.UtcNow;
        public int UploadBatch { get; set; }
    }
}

