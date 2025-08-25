using System.ComponentModel.DataAnnotations.Schema;
using System.ComponentModel.DataAnnotations;

namespace Tools.Models
{
    public class ProjectConfig
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public int Id { get; set; }
        public List<int> Modules { get; set; }
        public int ProjectId { get; set; }
        public string Envelope { get; set; }
        public List<int> BoxBreaking { get; set; }

    }
}
