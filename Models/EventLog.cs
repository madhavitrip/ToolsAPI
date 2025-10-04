using System.ComponentModel.DataAnnotations.Schema;
using System.ComponentModel.DataAnnotations;

namespace Tools.Models
{
    public class EventLog
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public int EventId { get; set; }
        public string Event { get; set; }
        public string Category { get; set; }
        public int EventTriggeredBy { get; set; }
        public DateTime LoggedAt { get; set; }
        public string OldValue { get; set; }
        public string NewValue { get; set; }
    }
}
