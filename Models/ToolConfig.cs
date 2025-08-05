namespace Tools.Models
{
    public class ToolConfig
    {
        public int ToolConfigId {  get; set; }
        public List<int> ToolId { get; set; }
        public int ProjectId { get; set; }
    }
}
