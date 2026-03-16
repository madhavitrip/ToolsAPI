namespace Tools.Models
{
    public class ConflictingFields
    {
        public int Id { get; set; }
        public int NRDataId { get; set; }
        public int ProjectId { get; set; }
        public string UniqueField { get; set; }
        public string ConflictingField { get; set; }
        public int? ChangedNRDataId {get;set;}
        public int Status { get; set; }
    }
}
