namespace Tools.Models
{
    public class Project
    {
        public int ProjectId { get; set; }
        public int GroupId { get; set; }
        public int TypeId { get; set; }
        public List<int> UserAssigned {  get; set; }
        public bool Status { get; set; }
    }
}
