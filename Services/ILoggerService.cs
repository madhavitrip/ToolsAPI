using System.Threading.Tasks;

namespace Tools.Services
{
    public interface ILoggerService
    {
        Task LogEventAsync(string message, string category, int triggeredBy, int ProjectId, string oldValue = null, string newValue = null);
        Task LogErrorAsync(string error, string errorMsg, string controller);
    }
}

