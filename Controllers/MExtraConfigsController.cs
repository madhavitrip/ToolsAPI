using ERPToolsAPI.Data;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Tools.Models;
using Tools.Services;

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
        public async Task<ActionResult<MExtraConfigurations>> PostMExtrasConfiguration(MExtraConfigurations extrasConfiguration)
        {
            try
            {
                var extra = await _context.MExtraConfigurations
                    .Where(x => x.TypeId == extrasConfiguration.TypeId
                             && x.GroupId == extrasConfiguration.GroupId
                             && x.ExtraType == extrasConfiguration.ExtraType)
                    .ToListAsync();

                if (extra.Any())
                {
                    var folderName = $"{extrasConfiguration.TypeId}_{extrasConfiguration.GroupId}";

                    var reportPath = Path.Combine(
                        Directory.GetCurrentDirectory(),
                        "wwwroot",
                        folderName
                    );

                    if (Directory.Exists(reportPath))
                    {
                        Directory.Delete(reportPath, true);
                    }

                    _context.MExtraConfigurations.RemoveRange(extra);
                    await _context.SaveChangesAsync();

                    _loggerService.LogEvent(
                        $"Deleted old MExtrasConfiguration record(s) for TypeId {extrasConfiguration.TypeId} and GroupId {extrasConfiguration.GroupId}",
                        "MExtraConfigurations",
                        User.Identity?.Name != null ? int.Parse(User.Identity.Name) : 0,
                        extrasConfiguration.GroupId
                    );
                }

                _context.MExtraConfigurations.Add(extrasConfiguration);
                await _context.SaveChangesAsync();

                _loggerService.LogEvent(
                    $"Created new MExtrasConfiguration with TypeId {extrasConfiguration.TypeId} and GroupId {extrasConfiguration.GroupId}",
                    "MExtraConfigurations",
                    User.Identity?.Name != null ? int.Parse(User.Identity.Name) : 0,
                    extrasConfiguration.GroupId
                );

                return Ok(extrasConfiguration); 
            }
            catch (DbUpdateConcurrencyException ex)
            {
                _loggerService.LogError(
                    "Concurrency error when saving MExtrasConfiguration",
                    ex.Message,
                    nameof(MExtraConfigsController)
                );

                return Conflict("Concurrency conflict occurred. The data may have been modified or deleted by another process.");
            }
            catch (Exception ex)
            {
                _loggerService.LogError(
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
