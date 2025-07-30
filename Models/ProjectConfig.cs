namespace Tools.Models
{
    public class ProjectConfig
    {
        public int Id { get; set; }
        public List<int> Modules { get; set; }
        public int ProjectId { get; set; }
        public List<int> MergeKey { get; set; }
        public List<int> PackagingSort { get; set; }
        public List<int> EnvelopeSort { get; set; }

    }
}
