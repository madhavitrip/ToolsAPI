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

            return Ok(projectConfig);
        }

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

        // POST: api/ProjectConfigs
        [HttpPost]
        public async Task<ActionResult<ProjectConfig>> PostProjectConfig(ProjectConfig projectConfig)
        {
            // Authorization check based on role
            int userRoleId = LogHelper.GetUserRoleId(User, Request);
            int userId = LogHelper.GetTriggeredBy(User);
            int projectId = projectConfig.ProjectId;

            // RoleId 1, 2, 3 (Developer, Admin, Head) have global access
            if (userRoleId == 1 || userRoleId == 2 || userRoleId == 3)
            {
                // Full access - proceed
            }
            // RoleId 4 (Manager) can only create for assigned projects
            else if (userRoleId == 4)
            {
                // Get the project to check if manager is assigned
                var project = await _context.Projects.FirstOrDefaultAsync(p => p.ProjectId == projectId);
                if (project == null || !project.UserAssigned.Contains(userId))
                {
                    await _loggerService.LogEventAsync(
                        $"AUTHORIZATION DENIED: Manager (RoleId 4, UserId {userId}) attempted to create ProjectConfig for ProjectId {projectId}",
                        "ProjectConfig",
                        userId,
                        projectId
                    );
                    return StatusCode(403, new { message = "You are not authorized to create configuration for this project." });
                }
            }
            else
            {
                await _loggerService.LogEventAsync(
                    $"AUTHORIZATION DENIED: User with RoleId {userRoleId} attempted to create ProjectConfig",
                    "ProjectConfig",
                    userId,
                    projectId
                );
                return StatusCode(403, new { message = "You are not authorized to create project configurations." });
            }

            try
            {
                _context.ProjectConfigs.Add(projectConfig);
                await _context.SaveChangesAsync();
                await _loggerService.LogEventAsync($"Created ProjectConfig for ProjectId {projectId}", "ProjectConfig", userId, projectId);
                return CreatedAtAction("GetProjectConfig", new { id = projectConfig.Id }, projectConfig);
            }
            catch (Exception ex)
            {
                await _loggerService.LogErrorAsync("Error creating ProjectConfigs", ex.Message, nameof(ProjectConfigsController));
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
    }
}
