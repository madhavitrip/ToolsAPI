namespace ToolsAPI.Models
{
    public class MasterAuthSettings
    {
        public string PasscodeHash { get; set; } = string.Empty;
        public int MaxFailedAttempts { get; set; } = 5;
        public int LockoutMinutes { get; set; } = 15;
    }
}
