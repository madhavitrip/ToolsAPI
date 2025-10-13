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
        public async Task<ActionResult<ProjectConfig>> GetProjectConfig(int id)
        {
            var projectConfig = await _context.ProjectConfigs.FindAsync(id);

            if (projectConfig == null)
            {
                return NotFound();
            }

            return projectConfig;
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
                _loggerService.LogEvent($"Updated ProjectConfig for {projectConfig.ProjectId}", "ProjectConfig", User.Identity?.Name != null ? int.Parse(User.Identity.Name) : 0, projectConfig.ProjectId);
            }
            catch (Exception ex)
            {
                if (!ProjectConfigExists(id))
                {
                    _loggerService.LogEvent($"ProjectConfig with ID {id} not found", "ProjectConfig", User.Identity?.Name != null ? int.Parse(User.Identity.Name) : 0, projectConfig.ProjectId);
                    return NotFound();
                }
                else
                {
                    _loggerService.LogError("Error updating ProjectConfigs", ex.Message, nameof(ProjectConfigsController));
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
                    _loggerService.LogEvent($"ProjectConfig for {config.ProjectId} already exists", "ProjectConfig", User.Identity?.Name != null ? int.Parse(User.Identity.Name) : 0, projectConfig.ProjectId);
                   _context.ProjectConfigs.Remove(config);
                    await _context.SaveChangesAsync();
                }
                _context.ProjectConfigs.Add(projectConfig);
                await _context.SaveChangesAsync();
                _loggerService.LogEvent($"Created a new ProjectConfig with ID {projectConfig.ProjectId}", "ProjectConfig", User.Identity?.Name != null ? int.Parse(User.Identity.Name) : 0, projectConfig.ProjectId);
                return CreatedAtAction("GetProjectConfig", new { id = projectConfig.Id }, projectConfig);
            }
            catch (Exception ex)
            {
              //  _loggerService.LogError("Error creating ProjectConfigs", ex.Message, nameof(ProjectConfigsController));
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

        private bool ProjectConfigExists(int id)
        {
            return _context.ProjectConfigs.Any(e => e.Id == id);
        }
    }
}
