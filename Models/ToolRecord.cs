namespace Tools.Models
{
    public class ToolRecord
    {
        public int Id { get; set; }
        public int UserId { get; set; } // from ERP JWT
        public string ToolName { get; set; }
        public string Description { get; set; }
        public DateTime CreatedAt { get; set; }
    }
}
