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
        public List<int> BoxBreakingCriteria { get; set; }
        public List<int> DuplicateRemoveFields { get; set; }
        public List<int> SortingBoxReport { get; set; }
        public List<int> EnvelopeMakingCriteria { get; set; }
        public int BoxCapacity { get; set; }
        public List<int> DuplicateCriteria { get; set; }
        public double Enhancement {  get; set; }
        public int BoxNumber { get; set; }
        public int OmrSerialNumber { get; set; }

    }
}
