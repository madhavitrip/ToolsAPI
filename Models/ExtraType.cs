using System.ComponentModel.DataAnnotations.Schema;
using System.ComponentModel.DataAnnotations;

namespace Tools.Models
{
    public class ExtraType
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public int ExtraTypeId { get; set; }
        public string Type {  get; set; }
        public int Order {  get; set; }
    }
}
