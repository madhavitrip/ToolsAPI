using ERPToolsAPI.Data;
using System;
using System.Threading;
using System.Threading.Tasks;
using Tools.Models;

namespace Tools.Services
{
    public class LoggerService : ILoggerService
    {
        private readonly ERPToolsDbContext _context;
        private static readonly SemaphoreSlim _lock = new SemaphoreSlim(1, 1);

        public LoggerService(ERPToolsDbContext context)
        {
            _context = context;
        }

        public async Task LogEventAsync(string message, string category, int triggeredBy, int ProjectId, string oldValue = null, string newValue = null)
        {
            await _lock.WaitAsync();
            try
            {
                var log = new EventLog
                {
                    Event = message,
                    EventTriggeredBy = triggeredBy,
                    ProjectId = ProjectId,
                    Category = category,
                    OldValue = oldValue,
                    NewValue = newValue
                };
                _context.EventLogs.Add(log);
                await _context.SaveChangesAsync();
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
                var log = new ErrorLog
                {
                    Error = error,
                    Message = errormessage,
                    Occurance = controller,
                };

                _context.ErrorLogs.Add(log);
                await _context.SaveChangesAsync();
            }
            finally
            {
                _lock.Release();
            }
        }
    }
}

