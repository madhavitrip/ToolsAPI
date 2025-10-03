using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using ERPToolsAPI.Data;
using Tools.Models;

namespace Tools.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ProjectConfigsController : ControllerBase
    {
        private readonly ERPToolsDbContext _context;

        public ProjectConfigsController(ERPToolsDbContext context)
        {
            _context = context;
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
            }
            catch (DbUpdateConcurrencyException)
            {
                if (!ProjectConfigExists(id))
                {
                    return NotFound();
                }
                else
                {
                    throw;
                }
            }

            return NoContent();
        }

        // POST: api/ProjectConfigs
        // To protect from overposting attacks, see https://go.microsoft.com/fwlink/?linkid=2123754
        [HttpPost]
        public async Task<ActionResult<ProjectConfig>> PostProjectConfig(ProjectConfig projectConfig)
        {
            var config = await _context.ProjectConfigs.Where(p=>p.ProjectId == projectConfig.ProjectId).FirstOrDefaultAsync();
            if (config!= null)
            {
                return Conflict(new { message = "A configuration already exists for this project." });
            }
            _context.ProjectConfigs.Add(projectConfig);
            await _context.SaveChangesAsync();

            return CreatedAtAction("GetProjectConfig", new { id = projectConfig.Id }, projectConfig);
        }

        // DELETE: api/ProjectConfigs/5
        [HttpDelete("{id}")]
        public async Task<IActionResult> DeleteProjectConfig(int id)
        {
            var projectConfig = await _context.ProjectConfigs.FindAsync(id);
            if (projectConfig == null)
            {
                return NotFound();
            }

            _context.ProjectConfigs.Remove(projectConfig);
            await _context.SaveChangesAsync();

            return NoContent();
        }

        private bool ProjectConfigExists(int id)
        {
            return _context.ProjectConfigs.Any(e => e.Id == id);
        }
    }
}
