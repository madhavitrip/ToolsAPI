using System.Threading.Tasks;

namespace Tools.Services
{
    public interface ILoggerService
    {
        Task LogEventAsync(string message, string category, int triggeredBy, int ProjectId, string? oldValue = null, string? newValue = null);
        Task LogErrorAsync(string error, string errorMsg, string controller);

        /// <summary>
        /// Logs a centralized API request/response audit entry.
        /// </summary>
        Task LogApiActivityAsync(
            string httpMethod,
            string endpoint,
            string category,
            int triggeredBy,
            int projectId,
            string? oldValue,
            string? newValue,
            int statusCode,
            string? details = null);

        /// <summary>
        /// Logs an entity or state change with strongly-typed old and new values.
        /// </summary>
        Task LogChangeAsync<T>(
            string message,
            string category,
            int triggeredBy,
            int projectId,
            T? oldValue,
            T? newValue);
    }
}
