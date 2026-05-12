using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using ERPToolsAPI.Data;
using Tools.Models;
using Tools.Services;
using Microsoft.CodeAnalysis;

namespace Tools.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ProjectConfigsController : ControllerBase
    {
        private readonly ERPToolsDbContext _context;
        private readonly ILoggerService _loggerService;
        public ProjectConfigsController(ERPToolsDbContext context, ILoggerService loggerService)
        {
            _context = context;
            _loggerService = loggerService;
        }

        // GET: api/ProjectConfigs
        [HttpGet]
        public async Task<ActionResult<IEnumerable<ProjectConfig>>> GetProjectConfigs()
        {
            return await _context.ProjectConfigs.ToListAsync();
        }

        // GET: api/ProjectConfigs/ByProject/77
        [HttpGet("ByProject/{projectId}")]
        public async Task<ActionResult<ProjectConfig>> GetProjectConfigByProjectId(int projectId)
        {
            var config = await _context.ProjectConfigs
                .FirstOrDefaultAsync(p => p.ProjectId == projectId);

            if (config == null)
            {
                return NotFound(new { message = $"No configuration found for ProjectId: {projectId}" });
            }

            return Ok(config);
        }


        // GET: api/ProjectConfigs/5
        [HttpGet("{id}")]
        public async Task<IActionResult> GetProjectConfig(int id)
        {
            var projectConfig = await _context.ProjectConfigs
                .FirstOrDefaultAsync(p => p.ProjectId == id);

            if (projectConfig == null)
                return NotFound();

            // Collect all field ids used in config
            var allFieldIds = new List<int>();

            void AddIds(List<int> ids)
            {
                if (ids != null && ids.Any())
                    allFieldIds.AddRange(ids);
            }

            AddIds(projectConfig.EnvelopeMakingCriteria);
            AddIds(projectConfig.BoxBreakingCriteria);
            AddIds(projectConfig.DuplicateCriteria);
            AddIds(projectConfig.DuplicateRemoveFields);
            AddIds(projectConfig.SortingBoxReport);
            AddIds(projectConfig.InnerBundlingCriteria);

            allFieldIds = allFieldIds.Distinct().ToList();

            // Fetch field names
            var fields = await _context.Fields
                .Where(f => allFieldIds.Contains(f.FieldId))
                .Select(f => new { f.FieldId, f.Name })
                .ToListAsync();

            // Helper function to convert ids -> names
            List<string> GetNames(List<int> ids)
            {
                if (ids == null) return new List<string>();

                return fields
                    .Where(f => ids.Contains(f.FieldId))
                    .Select(f => f.Name)
                    .ToList();
            }

            var result = new
            {
                projectConfig.Envelope,
                projectConfig.Modules,
                projectConfig.BoxCapacity,
                projectConfig.Enhancement,
                projectConfig.RoundOffBeforeEnhancement,
                projectConfig.BoxNumber,
                projectConfig.OmrSerialNumber,
                projectConfig.BookletSerialNumber,
                projectConfig.IsInnerBundlingDone,
                projectConfig.ResetOnSymbolChange,
                projectConfig.ResetOmrSerialOnCatchChange,
                projectConfig.ResetBookletSerialOnCatchChange,
                projectConfig.MssTypes,
                projectConfig.MssAttached,

                EnvelopeMakingCriteria = new
                {
                    Ids = projectConfig.EnvelopeMakingCriteria,
                    Names = GetNames(projectConfig.EnvelopeMakingCriteria)
                },

                BoxBreakingCriteria = new
                {
                    Ids = projectConfig.BoxBreakingCriteria,
                    Names = GetNames(projectConfig.BoxBreakingCriteria)
                },

                DuplicateCriteria = new
                {
                    Ids = projectConfig.DuplicateCriteria,
                    Names = GetNames(projectConfig.DuplicateCriteria)
                },

                DuplicateRemoveFields = new
                {
                    Ids = projectConfig.DuplicateRemoveFields,
                    Names = GetNames(projectConfig.DuplicateRemoveFields)
                },

                SortingBoxReport = new
                {
                    Ids = projectConfig.SortingBoxReport,
                    Names = GetNames(projectConfig.SortingBoxReport)
                },

                InnerBundlingCriteria = new
                {
                    Ids = projectConfig.InnerBundlingCriteria,
                    Names = GetNames(projectConfig.InnerBundlingCriteria)
                }
            };

            return Ok(result);
        }
        // PUT: api/ProjectConfigs/5
        // To protect from overposting attacks, see https://go.microsoft.com/fwlink/?linkid=2123754
        [HttpPut("{id}")]
        public async Task<IActionResult> PutProjectConfig(int id, ProjectConfig projectConfig)
        {
            if (id != projectConfig.Id)
            {
                return BadRequest();
            }

            _context.Entry(projectConfig).State = EntityState.Modified;

            try
            {
                await _context.SaveChangesAsync();
                await _loggerService.LogEventAsync($"Updated ProjectConfig for {projectConfig.ProjectId}", "ProjectConfig", LogHelper.GetTriggeredBy(User), projectConfig.ProjectId);
            }
            catch (Exception ex)
            {
                if (!ProjectConfigExists(id))
                {
                    await _loggerService.LogEventAsync($"ProjectConfig with ID {id} not found", "ProjectConfig", LogHelper.GetTriggeredBy(User), projectConfig.ProjectId);
                    return NotFound();
                }
                else
                {
                    await _loggerService.LogErrorAsync("Error updating ProjectConfigs", ex.Message, nameof(ProjectConfigsController));
                    return StatusCode(500, "Internal server error");

                }
            }

            return NoContent();
        }

        // POST: api/ProjectConfigs
        // To protect from overposting attacks, see https://go.microsoft.com/fwlink/?linkid=2123754
        [HttpPost]
        public async Task<ActionResult<ProjectConfig>> PostProjectConfig(ProjectConfig projectConfig)
        {
            try
            {
                var config = await _context.ProjectConfigs.Where(p => p.ProjectId == projectConfig.ProjectId).FirstOrDefaultAsync();
                
                if (config != null)
                {
                    await _loggerService.LogEventAsync($"ProjectConfig for {config.ProjectId} already exists", "ProjectConfig", LogHelper.GetTriggeredBy(User), projectConfig.ProjectId);
                   
                    _context.ProjectConfigs.Remove(config);
                    await _context.SaveChangesAsync();
                }
                _context.ProjectConfigs.Add(projectConfig);
                await _context.SaveChangesAsync();
                await _loggerService.LogEventAsync($"Created a new ProjectConfig with ID {projectConfig.ProjectId}", "ProjectConfig", LogHelper.GetTriggeredBy(User), projectConfig.ProjectId);
                return CreatedAtAction("GetProjectConfig", new { id = projectConfig.Id }, projectConfig);
            }
            catch (Exception ex)
            {
                await _loggerService.LogErrorAsync("Error creating ProjectConfigs", ex.Message, nameof(ProjectConfigsController));
                return StatusCode(500, "Internal server error");
            }
        }

        public class ModuleDeleteRequest
        {
            public int ProjectId { get; set; }
            public List<int> ModuleIds { get; set; } = new List<int>();
        }

        [HttpPost("DeleteModuleReports")]
        public async Task<IActionResult> DeleteModuleReports([FromBody] ModuleDeleteRequest request)
        {
            try
            {
                var userId = LogHelper.GetTriggeredBy(User);

                // ?? Step 1: Get module names from DB
                var moduleNames = await _context.Modules
                    .Where(m => request.ModuleIds.Contains(m.Id))
                    .Select(m => m.Name)
                    .ToListAsync();

                // ?? Step 2: Mapping (name ? report key)
                var moduleToReportKeyMap = GetModuleToReportKeyMap();

                var reportKeys = moduleNames
                    .Select(name => name?.Trim().ToLower())
                    .Where(name => name != null && moduleToReportKeyMap.ContainsKey(name))
                    .Select(name => moduleToReportKeyMap[name])
                    .Distinct()
                    .ToList();

                // ?? Step 3: Base path
                var basePath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", request.ProjectId.ToString());

                if (!Directory.Exists(basePath))
                    return Ok("No project folder found");

                // ?? Step 4: Delete matching files
                foreach (var key in reportKeys)
                {
                    var files = Directory.GetFiles(basePath, $"{key}*");

                    foreach (var file in files)
                    {
                        System.IO.File.Delete(file);
                    }
                }

                await _loggerService.LogEventAsync(
                    $"Deleted reports: {string.Join(",", reportKeys)}",
                    "ModuleCleanup",
                    userId,
                    request.ProjectId
                );

                return Ok(new
                {
                    DeletedModules = moduleNames,
                    DeletedReports = reportKeys
                });
            }
            catch (Exception ex)
            {
                await _loggerService.LogErrorAsync("Error deleting reports", ex.Message, nameof(ProjectConfigsController));
                return StatusCode(500, "Internal server error");
            }
        }
        private Dictionary<string, string> GetModuleToReportKeyMap()
        {
            return new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
    {
        { "duplicate tool", "duplicate" },
        { "extra configuration", "extra" },
        { "envelope breaking", "envelope" },
        { "box breaking", "box" },
        { "envelope summary", "envelopeSummary" },
        { "catch summary report", "catchSummary" },
        { "catchomrserialingreport", "catchOmrSerialing" }
    };
        }
        // DELETE: api/ProjectConfigs/5
        [HttpDelete("{id}")]
        public async Task<IActionResult> DeleteProjectConfig(int id)
        {
            try
            {
                var projectConfig = await _context.ProjectConfigs.FindAsync(id);
                if (projectConfig == null)
                {
                    await _loggerService.LogEventAsync($"ProjectConfig with ID {id} not found during delete", "ProjectConfig", LogHelper.GetTriggeredBy(User), projectConfig.ProjectId);
                    return NotFound();
                }

                _context.ProjectConfigs.Remove(projectConfig);
                await _context.SaveChangesAsync();
                await _loggerService.LogEventAsync($"Deleted a ProjectConfig with ID {projectConfig.ProjectId}", "ProjectConfig", LogHelper.GetTriggeredBy(User), projectConfig.ProjectId);
                return NoContent();
            }
            catch (Exception ex)
            {
                await _loggerService.LogErrorAsync("Error deleting ProjectConfigs", ex.Message, nameof(ProjectConfigsController));
                return StatusCode(500, "Internal server error");
            }
        }

        private bool ProjectConfigExists(int id)
        {
            return _context.ProjectConfigs.Any(e => e.Id == id);
        }
    }
}

