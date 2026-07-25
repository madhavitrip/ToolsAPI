using System.ComponentModel.DataAnnotations.Schema;
using System.ComponentModel.DataAnnotations;

namespace Tools.Models
{
    public class ProcessSteps
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public int ProcessId { get; set; }
        public string ProcessName { get; set; }
        public int Steps { get; set; }
    }
}
