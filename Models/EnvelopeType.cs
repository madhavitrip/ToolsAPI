using System.ComponentModel.DataAnnotations.Schema;
using System.ComponentModel.DataAnnotations;

namespace Tools.Models
{
    public class EnvelopeType
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public int EnvelopeId { get; set; }
        public string EnvelopeName { get; set; }
        public int Capacity { get; set; }

    }
}
