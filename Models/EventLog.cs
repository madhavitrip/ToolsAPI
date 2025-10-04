namespace Tools.Models
{
    public class EventLog
    {
        public int EventId { get; set; }
        public string Event { get; set; }
        public string Category { get; set; }
        public int EventTriggeredBy { get; set; }
        public DateTime LoggedAt { get; set; }
        public string OldValue { get; set; }
        public string NewValue { get; set; }
    }
}
