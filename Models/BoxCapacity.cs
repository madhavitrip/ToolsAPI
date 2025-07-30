using System.ComponentModel.DataAnnotations.Schema;
using System.ComponentModel.DataAnnotations;

namespace Tools.Models
{
    public class BoxCapacity
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public int BoxCapacityId { get; set; }
        public int Capacity { get; set; }
    }
}
