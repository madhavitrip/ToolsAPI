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

        //GET: api/ProjectConfigs/ByProject/77
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


        [HttpGet("{projectId}")]
        public async Task<IActionResult> GetProjectConfigByProjectsId(int projectId)
        {
            var config = await _context.ProjectConfigs
                .FirstOrDefaultAsync(p => p.ProjectId == projectId);

            if (config == null)
            {
                return NotFound(new { message = $"No configuration found for ProjectId: {projectId}" });
            }

            // Fetch all fields once
            var fields = await _context.Fields.ToListAsync();

            List<FieldDto> MapFields(List<int> ids)
            {
                return fields
                    .Where(f => ids.Contains(f.FieldId))
                    .Select(f => new FieldDto
                    {
                        FieldId = f.FieldId,
                        Name = f.Name
                    })
                    .ToList();
            }

            var result = new
            {
                config.Id,
                config.ProjectId,
                config.Envelope,

                config.Modules,

                config.BoxCapacity,
                config.Enhancement,
                config.BoxNumber,
                config.OmrSerialNumber,
                config.ResetOnSymbolChange,
                config.IsInnerBundlingDone,
                config.ResetOmrSerialOnCatchChange,
                config.BookletSerialNumber,
                config.ResetBookletSerialOnCatchChange,
                config.MssTypes,
                config.MssAttached,
                config.RoundOffBeforeEnhancement,

                DuplicateRemoveFields = MapFields(config.DuplicateRemoveFields),
                SortingBoxReport = MapFields(config.SortingBoxReport),
                EnvelopeMakingCriteria = MapFields(config.EnvelopeMakingCriteria),
                BoxBreakingCriteria = MapFields(config.BoxBreakingCriteria),
                DuplicateCriteria = MapFields(config.DuplicateCriteria),
                InnerBundlingCriteria = MapFields(config.InnerBundlingCriteria)
            };

            return Ok(result);
        }
        public class FieldDto
        {
            public int FieldId { get; set; }
            public string Name { get; set; }
        }

        // GET: api/ProjectConfigs/5
        //[HttpGet("{id}")]
        //public async Task<IActionResult> GetProjectConfig(int id)
        //{
        //    var projectConfig = await _context.ProjectConfigs
        //        .FirstOrDefaultAsync(p => p.ProjectId == id);

        //    if (projectConfig == null)
        //        return NotFound();

        //    return Ok(projectConfig);
        //}

        // PUT: api/ProjectConfigs/5
        [HttpPut("{id}")]
        public async Task<IActionResult> PutProjectConfig(int id, ProjectConfig projectConfig)
        {
            if (id != projectConfig.Id)
            {
                return BadRequest();
            }

            // Authorization check based on role
            int userRoleId = LogHelper.GetUserRoleId(User, Request);
            int userId = LogHelper.GetTriggeredBy(User);
            int projectId = projectConfig.ProjectId;

            // RoleId 1, 2, 3 (Developer, Admin, Head) have global access
            if (userRoleId == 1 || userRoleId == 2 || userRoleId == 3)
            {
                // Full access - proceed
            }
            // RoleId 4 (Manager) can only edit assigned projects
            else if (userRoleId == 4)
            {
                // Get the project to check if manager is assigned
                var project = await _context.Projects.FirstOrDefaultAsync(p => p.ProjectId == projectId);
                if (project == null || !project.UserAssigned.Contains(userId))
                {
                    await _loggerService.LogEventAsync(
                        $"AUTHORIZATION DENIED: Manager (RoleId 4, UserId {userId}) attempted to edit ProjectConfig for ProjectId {projectId}",
                        "ProjectConfig",
                        userId,
                        projectId
                    );
                    return StatusCode(403, new { message = "You are not authorized to edit configuration for this project." });
                }
            }
            else
            {
                await _loggerService.LogEventAsync(
                    $"AUTHORIZATION DENIED: User with RoleId {userRoleId} attempted to edit ProjectConfig",
                    "ProjectConfig",
                    userId,
                    projectId
                );
                return StatusCode(403, new { message = "You are not authorized to edit project configurations." });
            }

            _context.Entry(projectConfig).State = EntityState.Modified;

            try
            {
                await _context.SaveChangesAsync();
                await _loggerService.LogEventAsync($"Updated ProjectConfig for ProjectId {projectId}", "ProjectConfig", userId, projectId);
            }
            catch (Exception ex)
            {
                if (!ProjectConfigExists(id))
                {
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
        [HttpPost]
        public async Task<ActionResult<ProjectConfig>> PostProjectConfig(ProjectConfig projectConfig)
        {
            try
            {
                var oldConfig = await _context.ProjectConfigs.AsNoTracking().Where(p => p.ProjectId == projectConfig.ProjectId).FirstOrDefaultAsync();
                int minResetStep = -1;
                var affectedModules = new List<string>();

                if (oldConfig != null)
                {
                    await _loggerService.LogEventAsync($"ProjectConfig for {oldConfig.ProjectId} already exists. Replacing config.", "ProjectConfig", LogHelper.GetTriggeredBy(User), projectConfig.ProjectId);

                    minResetStep = GetMinResetStep(oldConfig, projectConfig);
                    affectedModules = GetAffectedModulesForStep(minResetStep);

                    // Remove existing configs first
                    var trackingConfigs = await _context.ProjectConfigs.Where(p => p.ProjectId == projectConfig.ProjectId).ToListAsync();
                    if (trackingConfigs.Any())
                    {
                        _context.ProjectConfigs.RemoveRange(trackingConfigs);
                        await _context.SaveChangesAsync();
                    }
                }

                projectConfig.Id = 0; // Ensure EF Core inserts a new record
                _context.ProjectConfigs.Add(projectConfig);
                await _context.SaveChangesAsync();

                if (oldConfig != null)
                {
                    if (minResetStep >= 0)
                    {
                        await ResetAndCleanAffectedReports(projectConfig.ProjectId, affectedModules);

                        await _context.NRDatas
                            .Where(x => x.ProjectId == projectConfig.ProjectId && x.Status == true && x.Steps > minResetStep)
                            .ExecuteUpdateAsync(s => s.SetProperty(x => x.Steps, minResetStep));
                    }
                }
                else
                {
                    // Brand new configuration, clear all template statuses just in case
                    await ResetReportStatus(projectConfig.ProjectId);
                }

                await _loggerService.LogEventAsync($"Created/Updated ProjectConfig with ID {projectConfig.ProjectId}", "ProjectConfig", LogHelper.GetTriggeredBy(User), projectConfig.ProjectId);
                return CreatedAtAction(nameof(GetProjectConfigByProjectId), new { projectId = projectConfig.ProjectId }, projectConfig);
            }
            catch (Exception ex)
            {
                await _loggerService.LogErrorAsync("Error creating/updating ProjectConfigs", ex.Message, nameof(ProjectConfigsController));
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
                    return NotFound();
                }

                // Authorization check based on role
                int userRoleId = LogHelper.GetUserRoleId(User, Request);
                int userId = LogHelper.GetTriggeredBy(User);
                int projectId = projectConfig.ProjectId;

                // RoleId 1, 2, 3 (Developer, Admin, Head) have global access
                if (userRoleId == 1 || userRoleId == 2 || userRoleId == 3)
                {
                    // Full access - proceed
                }
                // RoleId 4 (Manager) can only delete for assigned projects
                else if (userRoleId == 4)
                {
                    // Get the project to check if manager is assigned
                    var project = await _context.Projects.FirstOrDefaultAsync(p => p.ProjectId == projectId);
                    if (project == null || !project.UserAssigned.Contains(userId))
                    {
                        await _loggerService.LogEventAsync(
                            $"AUTHORIZATION DENIED: Manager (RoleId 4, UserId {userId}) attempted to delete ProjectConfig for ProjectId {projectId}",
                            "ProjectConfig",
                            userId,
                            projectId
                        );
                        return StatusCode(403, new { message = "You are not authorized to delete configuration for this project." });
                    }
                }
                else
                {
                    await _loggerService.LogEventAsync(
                        $"AUTHORIZATION DENIED: User with RoleId {userRoleId} attempted to delete ProjectConfig",
                        "ProjectConfig",
                        userId,
                        projectId
                    );
                    return StatusCode(403, new { message = "You are not authorized to delete project configurations." });
                }

                _context.ProjectConfigs.Remove(projectConfig);
                await _context.SaveChangesAsync();
                await _loggerService.LogEventAsync($"Deleted ProjectConfig for ProjectId {projectId}", "ProjectConfig", userId, projectId);
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

        private bool AreIntListsEqual(List<int>? list1, List<int>? list2)
        {
            if (list1 == null && list2 == null) return true;
            if (list1 == null || list2 == null) return false;
            return !list1.Except(list2).Any() && !list2.Except(list1).Any();
        }

        private List<string> GetAffectedModulesForStep(int minResetStep)
        {
            var affected = new List<string>();

            switch (minResetStep)
            {
                case 0: // STEP_UPLOADED - all subsequent modules affected
                    affected.Add("Duplicate Tool");
                    affected.Add("Envelope Setup and Enhancement");
                    affected.Add("Envelope Breaking");
                    affected.Add("Extra Configuration");
                    affected.Add("Box Breaking");
                    break;
                case 1: // STEP_DUP_PARTIAL
                    affected.Add("Envelope Setup and Enhancement");
                    affected.Add("Envelope Breaking");
                    affected.Add("Extra Configuration");
                    affected.Add("Box Breaking");
                    break;
                case 2: // STEP_ENHANCEMENT
                    affected.Add("Envelope Breaking");
                    affected.Add("Extra Configuration");
                    affected.Add("Box Breaking");
                    break;
                case 3: // STEP_ENV_BREAKING
                    affected.Add("Extra Configuration");
                    affected.Add("Box Breaking");
                    break;
                case 4: // STEP_AWAITING_EXTRA
                    affected.Add("Box Breaking");
                    break;
                case 5: // STEP_AWAITING_ENV
                    affected.Add("Box Breaking");
                    break;
            }

            return affected;
        }

        private int GetMinResetStep(ProjectConfig oldConfig, ProjectConfig newConfig)
        {
            int resetStep = -1;

            // Check for module-specific changes and determine reset step
            // Module 1 = Duplicate Tool (Step 0)
            if (!AreIntListsEqual(oldConfig.DuplicateCriteria, newConfig.DuplicateCriteria) ||
                !AreIntListsEqual(oldConfig.DuplicateRemoveFields, newConfig.DuplicateRemoveFields))
            {
                resetStep = 0; // STEP_UPLOADED
            }

            // Module 2 = Enhancement / Envelope Setup (Step 1)
            if (oldConfig.Enhancement != newConfig.Enhancement ||
                oldConfig.RoundOffBeforeEnhancement != newConfig.RoundOffBeforeEnhancement ||
                oldConfig.Envelope != newConfig.Envelope)
            {
                resetStep = Math.Max(resetStep, 1); // STEP_DUP_PARTIAL
            }

            if (oldConfig.Envelope != newConfig.Envelope)
            {
                resetStep = Math.Max(resetStep, 2); // STEP_DUP_PARTIAL
            }

            // Module 3 = Envelope Breaking / Making (Step 2)
            if (!AreIntListsEqual(oldConfig.EnvelopeMakingCriteria, newConfig.EnvelopeMakingCriteria) ||
                oldConfig.ResetOnSymbolChange != newConfig.ResetOnSymbolChange ||
                !AreIntListsEqual(oldConfig.MssTypes, newConfig.MssTypes) ||
                oldConfig.MssAttached != newConfig.MssAttached ||
                oldConfig.OmrSerialNumber != newConfig.OmrSerialNumber ||
                oldConfig.BookletSerialNumber != newConfig.BookletSerialNumber ||
                oldConfig.ResetOmrSerialOnCatchChange != newConfig.ResetOmrSerialOnCatchChange ||
                oldConfig.ResetBookletSerialOnCatchChange != newConfig.ResetBookletSerialOnCatchChange)
            {
                resetStep = Math.Max(resetStep, 4); // STEP_ENHANCEMENT
            }

            // Module 4 = Extra Configuration (Step 4)
            // No specific config fields for this module in current ProjectConfig model
            // If module toggle changes, check if it was disabled (removed from Modules list)
            if (!AreIntListsEqual(oldConfig.Modules, newConfig.Modules))
            {
                var oldHasModule4 = oldConfig.Modules?.Contains(4) ?? false;
                var newHasModule4 = newConfig.Modules?.Contains(4) ?? false;

                if (oldHasModule4 && !newHasModule4)
                {
                    // Module 4 was disabled - no step reversion needed
                }
                else if (!oldHasModule4 && newHasModule4)
                {
                    // Module 4 was enabled - data can proceed as-is
                }
                else if (oldHasModule4 && newHasModule4)
                {
                    // Module 4 remains enabled but other modules may have changed
                    // Handle module re-ordering or duplication changes
                    resetStep = Math.Max(resetStep, 3); // STEP_AWAITING_EXTRA
                }
            }

            // Module 5 = Box Breaking (Step 5)
            if (oldConfig.BoxCapacity != newConfig.BoxCapacity ||
                oldConfig.BoxNumber != newConfig.BoxNumber ||
                !AreIntListsEqual(oldConfig.BoxBreakingCriteria, newConfig.BoxBreakingCriteria) ||
                !AreIntListsEqual(oldConfig.SortingBoxReport, newConfig.SortingBoxReport) ||
                oldConfig.IsInnerBundlingDone != newConfig.IsInnerBundlingDone ||
                !AreIntListsEqual(oldConfig.InnerBundlingCriteria, newConfig.InnerBundlingCriteria))
            {
                resetStep = Math.Max(resetStep, 5); // STEP_AWAITING_ENV
            }

            return resetStep;
        }

        private async Task ResetAndCleanAffectedReports(int projectId, List<string> affectedModules)
        {
            try
            {
                // 1. Reset template statuses in database
                var dbModules = await _context.Modules.ToListAsync();
                var affectedModuleIds = dbModules
                    .Where(m => affectedModules.Contains(m.Name, StringComparer.OrdinalIgnoreCase))
                    .Select(m => m.Id)
                    .ToList();

                if (affectedModuleIds.Any())
                {
                    // Fetch templates for project
                    var templates = await _context.RPTTemplates
                        .Where(t => t.ProjectId == projectId && t.ReportStatus)
                        .ToListAsync();

                    var templatesToReset = templates
                        .Where(t => t.ModuleIds != null && t.ModuleIds.Any(id => affectedModuleIds.Contains(id)))
                        .ToList();

                    if (templatesToReset.Any())
                    {
                        foreach (var t in templatesToReset)
                        {
                            t.ReportStatus = false;
                        }
                        await _context.SaveChangesAsync();
                    }
                }

                // 2. Clean physical report files
                var moduleToReportKeyMap = GetModuleToReportKeyMap();
                var reportKeys = affectedModules
                    .Select(name => name?.Trim().ToLower())
                    .Where(name => name != null && moduleToReportKeyMap.ContainsKey(name))
                    .Select(name => moduleToReportKeyMap[name])
                    .Distinct()
                    .ToList();

                // If "Envelope Breaking" is affected, we should also delete envelopeSummary
                if (affectedModules.Contains("Envelope Breaking", StringComparer.OrdinalIgnoreCase))
                {
                    reportKeys.Add("envelopeSummary");
                }
                // If "Box Breaking" is affected, we should also delete catchSummary and catchOmrSerialing
                if (affectedModules.Contains("Box Breaking", StringComparer.OrdinalIgnoreCase))
                {
                    reportKeys.Add("catchSummary");
                    reportKeys.Add("catchOmrSerialing");
                }

                reportKeys = reportKeys.Distinct().ToList();

                var basePath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", projectId.ToString());
                if (Directory.Exists(basePath))
                {
                    foreach (var key in reportKeys)
                    {
                        var files = Directory.GetFiles(basePath, $"{key}*");
                        foreach (var file in files)
                        {
                            try
                            {
                                System.IO.File.Delete(file);
                            }
                            catch (Exception ex)
                            {
                                await _loggerService.LogErrorAsync($"Error deleting file: {file}", ex.Message, nameof(ProjectConfigsController));
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                await _loggerService.LogErrorAsync("Error resetting affected reports", ex.Message, nameof(ProjectConfigsController));
            }
        }

        private async Task ResetReportStatus(int projectId)
        {
            try
            {
                await _context.RPTTemplates
                    .Where(t => t.ProjectId == projectId && t.ReportStatus)
                    .ExecuteUpdateAsync(setters => setters.SetProperty(t => t.ReportStatus, false));
            }
            catch (Exception ex)
            {
                await _loggerService.LogErrorAsync("Report Status Reset Error", ex.Message, nameof(ProjectConfigsController));
            }
        }
    }
}

