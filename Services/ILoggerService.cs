namespace Tools.Services
{
    public interface ILoggerService
    {
        void LogEvent(string message, string category, int triggeredBy, string oldValue = null, string newValue = null);
        void LogError(string error, string errorMsg, string controller);
    }
}
