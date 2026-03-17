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
                projectConfig.BoxNumber,
                projectConfig.OmrSerialNumber,
                projectConfig.BookletSerialNumber,
                projectConfig.IsInnerBundlingDone,
                projectConfig.ResetOnSymbolChange,
                projectConfig.ResetOmrSerialOnCatchChange,
                projectConfig.ResetBookletSerialOnCatchChange,

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

            var existingProjectConfig = await _context.ProjectConfigs
                .AsNoTracking()
                .FirstOrDefaultAsync(p => p.Id == id);

            if (existingProjectConfig == null)
            {
                var triggeredBy = LogHelper.GetTriggeredBy(User);
                _loggerService.LogEvent(
                    $"ProjectConfig with ID {id} not found",
                    "ProjectConfig",
                    triggeredBy,
                    projectConfig.ProjectId,
                    LogHelper.ToJson(existingProjectConfig),
                    LogHelper.ToJson(projectConfig)
                );
                return NotFound();
            }

            _context.Entry(projectConfig).State = EntityState.Modified;

            try
            {
                await _context.SaveChangesAsync();
                var triggeredBy = LogHelper.GetTriggeredBy(User);
                _loggerService.LogEvent(
                    $"Updated ProjectConfig for {projectConfig.ProjectId}",
                    "ProjectConfig",
                    triggeredBy,
                    projectConfig.ProjectId,
                    LogHelper.ToJson(existingProjectConfig),
                    LogHelper.ToJson(projectConfig)
                );
            }
            catch (Exception ex)
            {
                _loggerService.LogError("Error updating ProjectConfigs", ex.Message, nameof(ProjectConfigsController));
                return StatusCode(500, "Internal server error");
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
                    var triggeredby = LogHelper.GetTriggeredBy(User);
                    _loggerService.LogEvent(
                        $"ProjectConfig for {config.ProjectId} already exists",
                        "ProjectConfig",
                        triggeredby,
                        projectConfig.ProjectId,
                        LogHelper.ToJson(config),
                        LogHelper.ToJson(projectConfig)
                    );
                    var reportPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", config.ProjectId.ToString());
                    if (Directory.Exists(reportPath))
                    {
                        Directory.Delete(reportPath, true); // 'true' allows recursive deletion of files and subdirectories
                    }
                    _context.ProjectConfigs.Remove(config);
                    await _context.SaveChangesAsync();
                }
                _context.ProjectConfigs.Add(projectConfig);
                await _context.SaveChangesAsync();
                var triggeredBy = LogHelper.GetTriggeredBy(User);
                _loggerService.LogEvent(
                    $"Created a new ProjectConfig with ID {projectConfig.ProjectId}",
                    "ProjectConfig",
                    triggeredBy,
                    projectConfig.ProjectId,
                    string.Empty,
                    LogHelper.ToJson(projectConfig)
                );
                return CreatedAtAction("GetProjectConfig", new { id = projectConfig.Id }, projectConfig);
            }
            catch (Exception ex)
            {
               _loggerService.LogError("Error creating ProjectConfigs", ex.Message, nameof(ProjectConfigsController));
                return StatusCode(500, "Internal server error");
            }
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
                    _loggerService.LogEvent($"ProjectConfig with ID {id} not found during delete", "ProjectConfig", User.Identity?.Name != null ? int.Parse(User.Identity.Name) : 0,projectConfig.ProjectId);
                    return NotFound();
                }

                _context.ProjectConfigs.Remove(projectConfig);
                await _context.SaveChangesAsync();
                _loggerService.LogEvent($"Deleted a ProjectConfig with ID {projectConfig.ProjectId}", "ProjectConfig", User.Identity?.Name != null ? int.Parse(User.Identity.Name) : 0, projectConfig.ProjectId);
                return NoContent();
            }
            catch (Exception ex)
            {
                _loggerService.LogError("Error deleting ProjectConfigs", ex.Message, nameof(ProjectConfigsController));
                return StatusCode(500, "Internal server error");
            }
        }

    }
}
