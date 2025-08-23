namespace Tools.Models
{
    public class ExtraEnvelopes
    {
        public int Id { get; set; }
        public int ProjectId { get; set; }
        public int NRDataId { get; set; }
        public int ExtraId { get; set; }
        public int Quantity { get; set; }
        public string InnerEnvelope {  get; set; }
        public string OuterEnvelope { get; set; }
    }
}
