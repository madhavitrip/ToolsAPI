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
        public string? CatchNo { get; set; }

        // Calculated fields from envelope breaking
        public int EnvQuantity { get; set; }
        public int CenterEnv { get; set; }
        public int TotalEnv { get; set; }
        public string? Env { get; set; }         // "1/2", "2/2"
        public int SerialNumber { get; set; }
        public string? BookletSerial { get; set; }

        // All fields needed for sorting in both EnvelopeBreakage and BoxBreaking
        public string? CenterCode { get; set; }
        public int CenterSort { get; set; }
        public string? ExamTime { get; set; }
        public string? ExamDate { get; set; }
        public int Quantity { get; set; }
        public string? NodalCode { get; set; }
        public double NodalSort { get; set; }
        public string? Route { get; set; }
        public int RouteSort { get; set; }
        public int NRQuantity { get; set; }
        public string? CourseName { get; set; }

        // Modified sort values (for extras with different sort values)
        public double? NodalSortModified { get; set; }
        public int? CenterSortModified { get; set; }
        public int? RouteSortModified { get; set; }
        public string? NodalCodeRef { get; set; }
        public string? RouteRef { get; set; }

        // Metadata
        public DateTime CreatedAt { get; set; } = DateTime.UtcNow;
        public int UploadBatch { get; set; }
    }
}





