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
        public async Task<ActionResult> GetProjects(int page = 1, int pageSize = 10)
        {
            if (page <= 0) page = 1;
            if (pageSize <= 0) pageSize = 10;

            var totalRecords = await _context.Projects.CountAsync();

            var projects = await _context.Projects
                .OrderByDescending(p => p.ProjectId) // Always order before Skip
                .Skip((page - 1) * pageSize)
                .Take(pageSize)
                .ToListAsync();

            var result = new
            {
                Page = page,
                PageSize = pageSize,
                TotalRecords = totalRecords,
                TotalPages = (int)Math.Ceiling(totalRecords / (double)pageSize),
                Data = projects
            };

            return Ok(result);
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
                var userIdClaim = jsonToken?.Claims.FirstOrDefault(c => c.Type == "userid");
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

                // Filter the projects where the userId is in the UserAssigned list AND Status is false (active)
                var userProjects = projects
                    .Where(p => p.UserAssigned.Contains(userIntId) && p.Status == false) // Perform client-side filtering
                    .ToList();

                if (userProjects == null || !userProjects.Any())
                {
                    return NotFound("No projects found for this user.");
                }

                var projectIds = userProjects.Select(p => p.ProjectId).ToList();

                // Perform the join and grouping
                var projectWithLastLoggedAt = await _context.EventLogs
                    .Where(e => projectIds.Contains(e.ProjectId))  // Filter EventLogs for the user's projects
                    .GroupBy(e => e.ProjectId)  // Group by ProjectId to get distinct projects
                    .Select(g => new
                    {
                        ProjectId = g.Key,
                        LatestLoggedAt = g.Max(e => e.LoggedAt),  // Get the most recent LoggedAt for each ProjectId
                    })
                    .ToListAsync();

                // Join the result with userProjects to get the project name and latest log time
                var result = userProjects
                 .Join(projectWithLastLoggedAt,
                  p => p.ProjectId,
                  l => l.ProjectId,
                    (p, l) => new
                    {
                      p.ProjectId,
                      p.GroupId,
                      p.TypeId,
                      LatestLoggedAt = l.LatestLoggedAt,
                      p.Status,
                    })
                 .OrderByDescending(x => x.LatestLoggedAt)   // ? Order by latest access
                  .Select(x => new
                  {
                   x.ProjectId,
                   x.GroupId,
                   x.TypeId,
                   LoggedAt = x.LatestLoggedAt,
                   TimeAgo = GetTimeAgo(x.LatestLoggedAt),
                   x.Status,
                  })
               .ToList();
                return Ok(result);
            }
            catch (Exception ex)
            {
                return BadRequest($"Error decoding token: {ex.Message}");
            }
        }

        [Authorize]
        [HttpGet("ArchivedProjects")]
        public async Task<ActionResult<IEnumerable<Project>>> GetArchivedProjectsByUser()
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

                // Extract the userId from the correct claim
                var userIdClaim = jsonToken?.Claims.FirstOrDefault(c => c.Type == "userid");
                var userId = userIdClaim?.Value;

                if (string.IsNullOrEmpty(userId))
                {
                    return Unauthorized("User ID not found in token.");
                }

                // Convert userId to integer
                if (!int.TryParse(userId, out int userIntId))
                {
                    return Unauthorized("Invalid User ID format.");
                }

                // Fetch all projects
                var projects = await _context.Projects
                    .ToListAsync();

                // Filter the projects where the userId is in the UserAssigned list AND Status is true (archived)
                var userArchivedProjects = projects
                    .Where(p => p.UserAssigned.Contains(userIntId) && p.Status == true)
                    .ToList();

                if (userArchivedProjects == null || !userArchivedProjects.Any())
                {
                    return Ok(new List<object>()); // Return empty list instead of 404
                }

                var projectIds = userArchivedProjects.Select(p => p.ProjectId).ToList();

                // Get last accessed time for archived projects
                var projectWithLastLoggedAt = await _context.EventLogs
                    .Where(e => projectIds.Contains(e.ProjectId))
                    .GroupBy(e => e.ProjectId)
                    .Select(g => new
                    {
                        ProjectId = g.Key,
                        LatestLoggedAt = g.Max(e => e.LoggedAt),
                    })
                    .ToListAsync();

                // Join and format result
                var result = userArchivedProjects
                 .Join(projectWithLastLoggedAt,
                  p => p.ProjectId,
                  l => l.ProjectId,
                    (p, l) => new
                    {
                      p.ProjectId,
                      p.GroupId,
                      p.TypeId,
                      LatestLoggedAt = l.LatestLoggedAt,
                      IsActive = !p.Status, // Status true = archived, so IsActive is false
                    })
                 .OrderByDescending(x => x.LatestLoggedAt)
                  .Select(x => new
                  {
                   x.ProjectId,
                   x.GroupId,
                   x.TypeId,
                   TimeAgo = GetTimeAgo(x.LatestLoggedAt),
                   x.IsActive,
                  })
               .ToList();
                
                return Ok(result);
            }
            catch (Exception ex)
            {
                _loggerService.LogError("Error fetching archived projects", ex.Message, nameof(ProjectsController));
                return BadRequest($"Error fetching archived projects: {ex.Message}");
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
                    _loggerService.LogEvent($"User with ID not found in token. ", "Projects", LogHelper.GetTriggeredBy(User),0);
                    return Unauthorized("User ID not found in token.");
                }

                // Convert userId to integer (if necessary)
                if (!int.TryParse(userId, out int userIntId))
                {
                    _loggerService.LogEvent($"Invalid User ID format. ", "Projects", LogHelper.GetTriggeredBy(User), 0);

                    return Unauthorized("Invalid User ID format.");
                }
                Console.WriteLine(userIntId);
                var recentProjects = _context.EventLogs
     .Where(s => s.EventTriggeredBy == userIntId)  // Filter by user
     .GroupBy(s => s.ProjectId)  // Group by ProjectId
     .Select(g => new
     {
         ProjectId = g.Key,
         LatestEventId = g.Max(s => s.EventId),  // Get the max EventId for each Project
         LatestLoggedAt = g.Max(s => s.LoggedAt),  // Get the most recent LoggedAt for each Project
     })
     .OrderByDescending(p => p.LatestEventId)  // Order by the latest EventId (desc)
     .Take(3)  // Take the top 3 results
     .ToList()
     .Select(p => new
     {
         p.ProjectId,
         p.LatestEventId,
         TimeAgo = GetTimeAgo(p.LatestLoggedAt)  // Call the helper method to format the time difference
     })
     .ToList();

                // Helper method to format time difference
              


                return Ok(recentProjects);
            }
            catch (Exception ex)
            {
                _loggerService.LogError("Error decoding token",ex.Message,nameof(ProjectsController));
                return BadRequest($"Error decoding token: {ex.Message}");
            }


        }

        private string GetTimeAgo(DateTime loggedAt)
        {
            var timeDifference = DateTime.Now - loggedAt;

            if (timeDifference.TotalDays >= 1)
            {
                return $"{(int)timeDifference.TotalDays} day{(timeDifference.TotalDays > 1 ? "s" : "")} ago";
            }
            if (timeDifference.TotalHours >= 1)
            {
                return $"{(int)timeDifference.TotalHours} hour{(timeDifference.TotalHours > 1 ? "s" : "")} ago";
            }
            if (timeDifference.TotalMinutes >= 1)
            {
                return $"{(int)timeDifference.TotalMinutes} minute{(timeDifference.TotalMinutes > 1 ? "s" : "")} ago";
            }
            return "Just now";  // If it's less than a minute ago
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

        [HttpGet("GetMssByType")]
        public async Task<IActionResult> GetMssByType(int typeId)
        {
            try
            {
                var getMss = await _context.Mss.Where(s => s.TypeId == typeId).ToListAsync();
                return Ok(getMss);
            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }
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
                _loggerService.LogEvent($"Updated Project with ID {project.ProjectId}", "Projects", LogHelper.GetTriggeredBy(User),project.ProjectId);
            }
            catch (Exception ex)
            {
                if (!ProjectExists(id))
                {
                    _loggerService.LogEvent($"Project with ID {id} not found during updating", "Projects", LogHelper.GetTriggeredBy(User), project.ProjectId);
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
                // Check if a project with the same ProjectId already exists
                bool exists = await _context.Projects.AnyAsync(p => p.ProjectId == project.ProjectId);
                if (exists)
                {
                    return Conflict($"Project with ID {project.ProjectId} already exists.");
                }

                _context.Projects.Add(project);
                await _context.SaveChangesAsync();

                _loggerService.LogEvent(
                    $"Created a new Project with ID {project.ProjectId}",
                    "Projects",
                    LogHelper.GetTriggeredBy(User),
                    project.ProjectId
                );
                Console.WriteLine($"GroupId: {project.GroupId}, TypeId: {project.TypeId}");
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
                    _loggerService.LogEvent($"Project with ID {id} not found during delete", "Projects", LogHelper.GetTriggeredBy(User), project.ProjectId);
                    return NotFound();
                }

                _context.Projects.Remove(project);
                await _context.SaveChangesAsync();
                _loggerService.LogEvent($"Deleted a Project with ID {project.ProjectId}", "Projects", LogHelper.GetTriggeredBy(User), project.ProjectId);
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

