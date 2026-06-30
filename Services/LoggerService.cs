using ERPToolsAPI.Data;
using Microsoft.EntityFrameworkCore;
using System;
using Tools.Models;

namespace Tools.Services
{
    public class LoggerService : ILoggerService
    {
        private readonly DbContextOptions<ERPToolsDbContext> _options;

        public LoggerService(DbContextOptions<ERPToolsDbContext> options)
        {
            _options = options;
        }

        public void LogEvent(string message, string category, int triggeredBy, int ProjectId, string oldValue = null, string newValue = null)
        {
            var log = new EventLog
            {
                Event = message,
                EventTriggeredBy = triggeredBy,
                ProjectId = ProjectId,
                Category = category,
                OldValue = oldValue,  // Log the old value if available
                NewValue = newValue   // Log the new value if available
            };

            using (var context = new ERPToolsDbContext(_options))
            {
                context.EventLogs.Add(log);
                context.SaveChanges();
            }
        }

        public void LogError(string error, string errormessage, string controller)
        {
            var log = new ErrorLog
            {
                Error = error,
                Message = errormessage,
                Occurance = controller,
            };

            using (var context = new ERPToolsDbContext(_options))
            {
                context.ErrorLogs.Add(log);
                context.SaveChanges();
            }
        }
    }
}

