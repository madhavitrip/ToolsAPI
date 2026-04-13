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
    public class MProjectConfigsController : ControllerBase
    {
        private readonly ERPToolsDbContext _context;
        private readonly ILoggerService _loggerService;
        public MProjectConfigsController(ERPToolsDbContext context, ILoggerService loggerService)
        {
            _context = context;
            _loggerService = loggerService;
        }

        // POST: api/MProjectConfigs
        [HttpPost]
        public async Task<ActionResult<MProjectConfigs>> PostMProjectConfig(MProjectConfigs projectConfig)
        {
            try
            {
                var config = await _context.MProjectConfigs
                    .Where(p => p.TypeId == projectConfig.TypeId && p.GroupId == projectConfig.GroupId)
                    .FirstOrDefaultAsync();

                if (config != null)
                {
                    _loggerService.LogEvent(
                        $"MProjectConfig for TypeId {config.TypeId} and GroupId {config.GroupId} already exists",
                        "MProjectConfigs",
                        User.Identity?.Name != null ? int.Parse(User.Identity.Name) : 0,
                        config.GroupId
                    );

                    var reportPath = Path.Combine(
                        Directory.GetCurrentDirectory(),
                        "wwwroot",
                        $"{config.TypeId}_{config.GroupId}"
                    );

                    if (Directory.Exists(reportPath))
                    {
                        Directory.Delete(reportPath, true);
                    }

                    _context.MProjectConfigs.Remove(config);
                    await _context.SaveChangesAsync();
                }

                _context.MProjectConfigs.Add(projectConfig);
                await _context.SaveChangesAsync();

                _loggerService.LogEvent(
                    $"Created a new MProjectConfig with TypeId {projectConfig.TypeId} and GroupId {projectConfig.GroupId}",
                    "MProjectConfigs",
                    User.Identity?.Name != null ? int.Parse(User.Identity.Name) : 0,
                    projectConfig.GroupId
                );

                return Ok(projectConfig);
            }
            catch (Exception ex)
            {
                _loggerService.LogError("Error creating MProjectConfigs", ex.Message, nameof(MProjectConfigsController));
                return StatusCode(500, "Internal server error");
            }
        }

        [HttpGet]
        public async Task<ActionResult<IEnumerable<int>>> GetAllGroups()
        {
            try
            {
                var groupIds = await _context.MProjectConfigs
                    .Select(x => x.GroupId)
                    .Distinct()
                    .ToListAsync();

                return Ok(groupIds);
            }
            catch (Exception ex)
            {
                return StatusCode(500, ex.Message);
            }
        }

        [HttpGet("SelectedGroup")]
        public async Task<ActionResult<IEnumerable<int>>> GetAllTypeOfGroup(int GroupId)
        {
            try
            {
                var groupIds = await _context.MProjectConfigs.Where(s=>s.GroupId == GroupId)
                    .Select(x => x.TypeId)
                    .Distinct()
                    .ToListAsync();

                return Ok(groupIds);
            }
            catch (Exception ex)
            {
                return StatusCode(500, ex.Message);
            }
        }

        [HttpGet("ByTypeGroup/{typeId}/{groupId}")]
        public async Task<ActionResult<MProjectConfigs>> GetMProjectConfigByTypeGroup(int typeId, int groupId)
        {
            var config = await _context.MProjectConfigs
                .FirstOrDefaultAsync(p => p.TypeId == typeId && p.GroupId == groupId);

            if (config == null)
            {
                return NotFound(new { message = $"No configuration found for TypeId: {typeId} and GroupId: {groupId}" });
            }

            return Ok(config);
        }

    }
}
