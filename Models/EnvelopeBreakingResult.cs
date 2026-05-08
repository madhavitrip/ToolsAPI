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
        public int NrDataId { get; set; }      // Always set - for extras, use the base NRData's NrDataId
        public int? ExtraId { get; set; }       // NULL if regular NRData (1=Nodal, 2=University, 3=Office)
        public string? CatchNo { get; set; }

        // Calculated fields from envelope breaking
        public string EnvQuantity { get; set; }
        public int CenterEnv { get; set; }
        public int TotalEnv { get; set; }
        public string? Env { get; set; }         // "1/2", "2/2"
        public int SerialNumber { get; set; }
        public string? BookletSerial { get; set; }
        public string? OmrSerial { get; set; }
        // All fields needed for sorting in both EnvelopeBreakage and BoxBreaking
        public string? CenterCode { get; set; }
        public double CenterSort { get; set; }
        public string? ExamTime { get; set; }
        public string? ExamDate { get; set; }
        public int Quantity { get; set; }
        public string? NodalCode { get; set; }
        public double NodalSort { get; set; }
        public string? Route { get; set; }
        public int RouteSort { get; set; }
        public string? CourseName { get; set; }
        public string? District { get; set; }
        public int DistrictSort { get; set; }
        public DateTime CreatedAt { get; set; } = DateTime.UtcNow;
        public int UploadBatch { get; set; }
        public bool Status { get; set; }
    }
}





