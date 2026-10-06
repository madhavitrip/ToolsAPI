using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace Tools.Models
{
    public class CenterList
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public int Id { get; set; }
        public int ProjectId { get; set; }
        public string? NodalCode { get; set; }
        public string? CenterCode { get; set; }
        public string? Route { get; set; }
        public int RouteSort { get; set; }
        public double CenterSort { get; set; }
        public double NodalSort { get; set; }
        public int NRDataId { get; set; }
        public string CenterData {  get; set; }
        public int Quantity { get; set; }
        public int NRQuantity { get; set; }
        public bool Status { get; set; }
        public string? District { get; set; }
        public int DistrictSort { get; set; }
        public int ExtraId { get; set; }
    }
}
