using System.ComponentModel.DataAnnotations.Schema;
using System.ComponentModel.DataAnnotations;

namespace Tools.Models
{
    public class ToolConfig
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public int ToolConfigId {  get; set; }
        public List<int> ToolId { get; set; }
        public int ProjectId { get; set; }
    }
}
