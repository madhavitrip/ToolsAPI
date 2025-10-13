using ERPToolsAPI.Data;
using System;
using Tools.Models;

namespace Tools.Services
{
    public class LoggerService : ILoggerService
    {
        private readonly ERPToolsDbContext _context;

        public LoggerService(ERPToolsDbContext context)
        {
            _context = context;
        }

        public void LogEvent(string message, string category, int triggeredBy,int ProjectId, string oldValue = null, string newValue = null)
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
            _context.EventLogs.Add(log);
            _context.SaveChanges();
        }

        public void LogError(string error, string errormessage, string controller)
        {
            var log = new ErrorLog
            {
                Error = error,
                Message = errormessage,
                Occurance = controller,
            };

            _context.ErrorLogs.Add(log);
            _context.SaveChanges();
        }
    }
}
