using System;
using System.Threading;
using System.Threading.Tasks;
using ERPToolsAPI.Data;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using Tools.Models;

namespace Tools.Services
{
    public class LoggerService : ILoggerService
    {
        private readonly IServiceScopeFactory _scopeFactory;
        private readonly ILogger<LoggerService> _systemLogger;
        private static readonly SemaphoreSlim _lock = new SemaphoreSlim(1, 1);

        public LoggerService(IServiceScopeFactory scopeFactory, ILogger<LoggerService> systemLogger)
        {
            _scopeFactory = scopeFactory;
            _systemLogger = systemLogger;
        }

        public async Task LogEventAsync(string message, string category, int triggeredBy, int ProjectId, string? oldValue = null, string? newValue = null)
        {
            await _lock.WaitAsync();
            try
            {
                using var scope = _scopeFactory.CreateScope();
                var context = scope.ServiceProvider.GetRequiredService<ERPToolsDbContext>();

                var log = new EventLog
                {
                    Event = message,
                    EventTriggeredBy = triggeredBy,
                    ProjectId = ProjectId,
                    Category = category,
                    OldValue = oldValue,
                    NewValue = newValue,
                    LoggedAt = DateTime.Now
                };

                context.EventLogs.Add(log);
                await context.SaveChangesAsync();
            }
            catch (Exception ex)
            {
                _systemLogger.LogError(ex, "Failed to save EventLog: {Message}", message);
            }
            finally
            {
                _lock.Release();
            }
        }

        public async Task LogErrorAsync(string error, string errormessage, string controller)
        {
            await _lock.WaitAsync();
            try
            {
                using var scope = _scopeFactory.CreateScope();
                var context = scope.ServiceProvider.GetRequiredService<ERPToolsDbContext>();

                var log = new ErrorLog
                {
                    Error = error,
                    Message = errormessage,
                    Occurance = controller,
                };

                context.ErrorLogs.Add(log);
                await context.SaveChangesAsync();
            }
            catch (Exception ex)
            {
                _systemLogger.LogError(ex, "Failed to save ErrorLog: {Error}", error);
            }
            finally
            {
                _lock.Release();
            }
        }

        public async Task LogApiActivityAsync(
            string httpMethod,
            string endpoint,
            string category,
            int triggeredBy,
            int projectId,
            string? oldValue,
            string? newValue,
            int statusCode,
            string? details = null)
        {
            var eventDescription = $"[{httpMethod.ToUpper()}] {endpoint} - Status {statusCode}";
            if (!string.IsNullOrWhiteSpace(details))
            {
                eventDescription += $" ({details})";
            }

            await LogEventAsync(
                message: eventDescription,
                category: string.IsNullOrWhiteSpace(category) ? "API" : category,
                triggeredBy: triggeredBy,
                ProjectId: projectId,
                oldValue: oldValue,
                newValue: newValue
            );
        }

        public async Task LogChangeAsync<T>(
            string message,
            string category,
            int triggeredBy,
            int projectId,
            T? oldValue,
            T? newValue)
        {
            string? oldJson = AuditJsonHelper.SerializeSafe(oldValue);
            string? newJson = AuditJsonHelper.SerializeSafe(newValue);

            await LogEventAsync(
                message: message,
                category: category,
                triggeredBy: triggeredBy,
                ProjectId: projectId,
                oldValue: oldJson,
                newValue: newJson
            );
        }
    }
}
