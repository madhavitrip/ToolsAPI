using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using ERPToolsAPI.Data;
using Tools.Models;
using Microsoft.IdentityModel.Tokens;
using System.IdentityModel.Tokens.Jwt;
using Microsoft.AspNetCore.Authorization;
using Tools.Services;
using System.Data;

namespace Tools.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ProjectsController : ControllerBase
    {
        private readonly ERPToolsDbContext _context;
        private readonly ILoggerService _loggerService;
        public ProjectsController(ERPToolsDbContext context, ILoggerService loggerService)
        {
            _context = context;
            _loggerService = loggerService;
        }

        // GET: api/Projects
        [HttpGet]
        public async Task<ActionResult<IEnumerable<Project>>> GetProjects()
        {
            return await _context.Projects.ToListAsync();
        }

        [Authorize]
        [HttpGet("UserId")]
        public async Task<ActionResult<IEnumerable<Project>>> GetProjectByUser()
        {
            // Extract token from the Authorization header
            var token = Request.Headers["Authorization"].ToString()?.Replace("Bearer ", "");

            if (string.IsNullOrEmpty(token))
            {
                return Unauthorized("Token is required.");
            }

            try
            {
                // Decode the JWT token and extract userId
                var handler = new JwtSecurityTokenHandler();
                var jsonToken = handler.ReadToken(token) as JwtSecurityToken;

                // Extract the userId from the correct claim (adjusting based on your token structure)
                var userIdClaim = jsonToken?.Claims.FirstOrDefault(c => c.Type == "http://schemas.xmlsoap.org/ws/2005/05/identity/claims/name");
                var userId = userIdClaim?.Value;

                if (string.IsNullOrEmpty(userId))
                {
                    return Unauthorized("User ID not found in token.");
                }

                // Convert userId to integer (if necessary)
                if (!int.TryParse(userId, out int userIntId))
                {
                    return Unauthorized("Invalid User ID format.");
                }

                // Fetch all projects (we'll filter them in memory)
                var projects = await _context.Projects
                    .ToListAsync(); // Fetch all projects (client-side filtering will follow)

                // Filter the projects where the userId is in the UserAssigned list
                var userProjects = projects
                    .Where(p => p.UserAssigned.Contains(userIntId)) // Perform client-side filtering
                    .ToList();

                if (userProjects == null || !userProjects.Any())
                {
                    return NotFound("No projects found for this user.");
                }

                return Ok(userProjects);
            }
            catch (Exception ex)
            {
                return BadRequest($"Error decoding token: {ex.Message}");
            }
        }

        [HttpGet("RecentProjects")]

        public async Task<ActionResult> GetRecentProject()
        {
            var token = Request.Headers["Authorization"].ToString()?.Replace("Bearer ", "");

            if (string.IsNullOrEmpty(token))
            {
                return Unauthorized("Token is required.");
            }

            try
            {
                // Decode the JWT token and extract userId
                var handler = new JwtSecurityTokenHandler();
                var jsonToken = handler.ReadToken(token) as JwtSecurityToken;

                // Extract the userId from the correct claim (adjusting based on your token structure)
                var userIdClaim = jsonToken?.Claims.FirstOrDefault(c => c.Type == "http://schemas.xmlsoap.org/ws/2005/05/identity/claims/name");
                var userId = userIdClaim?.Value;

                if (string.IsNullOrEmpty(userId))
                {
                    return Unauthorized("User ID not found in token.");
                }

                // Convert userId to integer (if necessary)
                if (!int.TryParse(userId, out int userIntId))
                {
                    return Unauthorized("Invalid User ID format.");
                }
                var RecentActivity = _context.EventLogs.Where(s => s.EventTriggeredBy == userIntId).Select(s => new { s.ProjectId, s.LoggedAt, s.EventId }).OrderByDescending(s => s.EventId).Distinct().Take(3);
                return Ok(RecentActivity);
            }
            catch (Exception ex)
            {
                _loggerService.LogError("Error decoding token",ex.Message,nameof(ProjectsController));
                return BadRequest($"Error decoding token: {ex.Message}");
            }


        }
        // GET: api/Projects/5
        [HttpGet("{id}")]
        public async Task<ActionResult<Project>> GetProject(int id)
        {
            var project = await _context.Projects.FindAsync(id);

            if (project == null)
            {
                return NotFound();
            }

            return project;
        }

        // PUT: api/Projects/5
        // To protect from overposting attacks, see https://go.microsoft.com/fwlink/?linkid=2123754
        [HttpPut("{id}")]
        public async Task<IActionResult> PutProject(int id, Project project)
        {
            if (id != project.ProjectId)
            {
                return BadRequest();
            }

            _context.Entry(project).State = EntityState.Modified;

            try
            {
                await _context.SaveChangesAsync();
                _loggerService.LogEvent($"Updated Project with ID {project.ProjectId}", "Projects", User.Identity?.Name != null ? int.Parse(User.Identity.Name) : 0,project.ProjectId);
            }
            catch (Exception ex)
            {
                if (!ProjectExists(id))
                {
                    _loggerService.LogEvent($"Project with ID {id} not found during updating", "Projects", User.Identity?.Name != null ? int.Parse(User.Identity.Name) : 0, project.ProjectId);
                    return NotFound();
                }
                else
                {
                    _loggerService.LogError("Error creating Project", ex.Message, nameof(ProjectsController));
                    throw;
                }
            }

            return NoContent();
        }

        // POST: api/Projects
        // To protect from overposting attacks, see https://go.microsoft.com/fwlink/?linkid=2123754
        [HttpPost]
        public async Task<ActionResult<Project>> PostProject(Project project)
        {
            try
            {
                _context.Projects.Add(project);
                await _context.SaveChangesAsync();
                _loggerService.LogEvent($"Created a new Project with ID {project.ProjectId}", "Projects", User.Identity?.Name != null ? int.Parse(User.Identity.Name) : 0, project.ProjectId);
                return CreatedAtAction("GetProject", new { id = project.ProjectId }, project);
            }
            catch (Exception ex)
            {
                _loggerService.LogError("Error creating Project", ex.Message, nameof(ProjectsController));
                return StatusCode(500, "Internal server error");
            }
        }

        // DELETE: api/Projects/5
        [HttpDelete("{id}")]
        public async Task<IActionResult> DeleteProject(int id)
        {
            try
            {
                var project = await _context.Projects.FindAsync(id);
                if (project == null)
                {
                    _loggerService.LogEvent($"Project with ID {id} not found during delete", "Projects", User.Identity?.Name != null ? int.Parse(User.Identity.Name) : 0, project.ProjectId);
                    return NotFound();
                }

                _context.Projects.Remove(project);
                await _context.SaveChangesAsync();
                _loggerService.LogEvent($"Deleted a Project with ID {project.ProjectId}", "Projects", User.Identity?.Name != null ? int.Parse(User.Identity.Name) : 0, project.ProjectId);
                return NoContent();
            }
            catch (Exception ex)
            {
                _loggerService.LogError("Error deleting Project", ex.Message, nameof(ProjectsController));
                return StatusCode(500, "Internal server error");
            }
        }

        private bool ProjectExists(int id)
        {
            return _context.Projects.Any(e => e.ProjectId == id);
        }
    }
}
