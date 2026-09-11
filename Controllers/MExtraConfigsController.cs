using ERPToolsAPI.Data;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Tools.Models;
using Tools.Services;
using Tools.Middleware;
using System.Text.Json;

namespace Tools.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class MExtraConfigsController : ControllerBase
    {
        private readonly ERPToolsDbContext _context;
        private readonly ILoggerService _loggerService;
        public MExtraConfigsController(ERPToolsDbContext context, ILoggerService loggerService)
        {
            _context = context;
            _loggerService = loggerService;
        }

        // POST: api/MExtrasConfigs
        [HttpPost]
        [RequireMasterAuth(Module = "MExtraConfigs", Operation = "SAVE MASTER")]
        public async Task<ActionResult<MExtraConfigurations>> PostMExtrasConfiguration(MExtraConfigurations extrasConfiguration)
        {
            try
            {
                var extra = await _context.MExtraConfigurations
                    .Where(x => x.TypeId == extrasConfiguration.TypeId
                             && x.GroupId == extrasConfiguration.GroupId
                             && x.ExtraType == extrasConfiguration.ExtraType)
                    .ToListAsync();

                string oldValue = null;
                if (extra.Any())
                {
                    // Serialize old values for logging
                    oldValue = JsonSerializer.Serialize(extra.Select(e => new
                    {
                        e.Id,
                        e.GroupId,
                        e.TypeId,
                        e.ExtraType,
                        e.Mode,
                        e.Value,
                        e.EnvelopeType
                    }));

                    var folderName = $"{extrasConfiguration.TypeId}_{extrasConfiguration.GroupId}";

                    var reportPath = FileStorageHelper.GetGroupFolder(folderName);
                    if (Directory.Exists(reportPath))
                    {
                        try { Directory.Delete(reportPath, true); } catch { }
                    }

                    var legacyReportPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", folderName);
                    if (Directory.Exists(legacyReportPath))
                    {
                        try { Directory.Delete(legacyReportPath, true); } catch { }
                    }

                    _context.MExtraConfigurations.RemoveRange(extra);
                    await _context.SaveChangesAsync();

                    await _loggerService.LogEventAsync(
                        $"Deleted old MExtrasConfiguration record(s) for TypeId {extrasConfiguration.TypeId} and GroupId {extrasConfiguration.GroupId}",
                        "MExtraConfigurations",
                        LogHelper.GetTriggeredBy(User),
                        extrasConfiguration.GroupId,
                        oldValue,
                        null
                    );
                }

                // Serialize new value for logging
                var newValue = JsonSerializer.Serialize(new
                {
                    extrasConfiguration.Id,
                    extrasConfiguration.GroupId,
                    extrasConfiguration.TypeId,
                    extrasConfiguration.ExtraType,
                    extrasConfiguration.Mode,
                    extrasConfiguration.Value,
                    extrasConfiguration.EnvelopeType
                });

                _context.MExtraConfigurations.Add(extrasConfiguration);
                await _context.SaveChangesAsync();

                await _loggerService.LogEventAsync(
                    $"Created new MExtrasConfiguration with TypeId {extrasConfiguration.TypeId} and GroupId {extrasConfiguration.GroupId}",
                    "MExtraConfigurations",
                    LogHelper.GetTriggeredBy(User),
                    extrasConfiguration.GroupId,
                    oldValue,
                    newValue
                );

                return Ok(extrasConfiguration); 
            }
            catch (DbUpdateConcurrencyException ex)
            {
                await _loggerService.LogErrorAsync(
                    "Concurrency error when saving MExtrasConfiguration",
                    ex.Message,
                    nameof(MExtraConfigsController)
                );

                return Conflict("Concurrency conflict occurred. The data may have been modified or deleted by another process.");
            }
            catch (Exception ex)
            {
                await _loggerService.LogErrorAsync(
                    "Error creating MExtrasConfiguration",
                    ex.Message,
                    nameof(MExtraConfigsController)
                );

                return StatusCode(500, "Internal server error");
            }
        }


        [HttpGet("ByTypeGroup/{typeId}/{groupId}")]
        public async Task<ActionResult<IEnumerable<MExtraConfigurations>>> GetExtrasByTypeGroup(int typeId, int groupId)
        {
            var extras = await _context.MExtraConfigurations
                .Where(e => e.TypeId == typeId && e.GroupId == groupId)
                .ToListAsync();

            if (extras == null || extras.Count == 0)
            {
                return NotFound(new { message = $"No configurations found for TypeId: {typeId} and GroupId: {groupId}" });
            }

            return Ok(extras);
        }
    }


}

