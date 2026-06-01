using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using ERPToolsAPI.Data;
using ToolsAPI.Models;
using System.ComponentModel.DataAnnotations;

namespace ToolsAPI.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class EnvelopeLotReportsController : ControllerBase
    {
        private readonly ERPToolsDbContext _context;

        public EnvelopeLotReportsController(ERPToolsDbContext context)
        {
            _context = context;
        }

        // GET: api/EnvelopeLotReports/Test
        [HttpGet("Test")]
        public IActionResult Test()
        {
            return Ok(new { message = "EnvelopeLotReports API is working", timestamp = DateTime.UtcNow });
        }

        // GET: api/EnvelopeLotReports/ByProject/{projectId}
        [HttpGet("ByProject/{projectId}")]
        public async Task<ActionResult<IEnumerable<EnvelopeLotReport>>> GetEnvelopeLotReportsByProject(int projectId)
        {
            try
            {
                Console.WriteLine($"Loading envelope lot reports for project: {projectId}");
                var reports = await _context.EnvelopeLotReports
                    .Where(r => r.ProjectId == projectId)
                    .OrderByDescending(r => r.GeneratedAt)
                    .ToListAsync();

                Console.WriteLine($"Found {reports.Count} reports for project {projectId}");
                return Ok(reports);
            }
            catch (Exception ex)
            {
                var fullMessage = ex.Message + (ex.InnerException != null ? " | Inner: " + ex.InnerException.Message : "");
                Console.WriteLine($"Error loading reports for project {projectId}: {fullMessage}");
                return StatusCode(500, new { message = "Failed to retrieve reports", error = fullMessage });
            }
        }

        // GET: api/EnvelopeLotReports/ByTemplate/{templateId}/{projectId}
        [HttpGet("ByTemplate/{templateId}/{projectId}")]
        public async Task<ActionResult<IEnumerable<EnvelopeLotReport>>> GetEnvelopeLotReportsByTemplate(int templateId, int projectId)
        {
            try
            {
                var reports = await _context.EnvelopeLotReports
                    .Where(r => r.TemplateId == templateId && r.ProjectId == projectId)
                    .OrderByDescending(r => r.GeneratedAt)
                    .ToListAsync();

                return Ok(reports);
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { message = "Failed to retrieve reports", error = ex.Message });
            }
        }

        // POST: api/EnvelopeLotReports
        [HttpPost]
        public async Task<ActionResult<EnvelopeLotReport>> CreateEnvelopeLotReport([FromBody] CreateEnvelopeLotReportRequest request)
        {
            try
            {
                Console.WriteLine($"Received request to create envelope lot report: ProjectId={request.ProjectId}, TemplateId={request.TemplateId}, EnvLotNumbers={request.EnvLotNumbers}");
                
                if (!ModelState.IsValid)
                {
                    Console.WriteLine("Model state is invalid:");
                    foreach (var error in ModelState)
                    {
                        Console.WriteLine($"  {error.Key}: {string.Join(", ", error.Value.Errors.Select(e => e.ErrorMessage))}");
                    }
                    return BadRequest(ModelState);
                }

                // Always create a new report record instead of overwriting existing ones
                // This allows for a full history of generated reports for the project/template
                Console.WriteLine("Creating new report record for history");
                
                var newReport = new EnvelopeLotReport
                {
                    ProjectId = request.ProjectId,
                    TemplateId = request.TemplateId,
                    TemplateName = request.TemplateName,
                    EnvLotNumbers = request.EnvLotNumbers ?? "",
                    FileName = request.FileName,
                    GeneratedAt = DateTime.UtcNow,
                    GeneratedBy = request.GeneratedBy,
                    FilePath = request.FilePath
                };

                _context.EnvelopeLotReports.Add(newReport);
                await _context.SaveChangesAsync();
                Console.WriteLine($"New report created with ID: {newReport.Id}");

                return CreatedAtAction(nameof(GetEnvelopeLotReportsByProject), 
                    new { projectId = newReport.ProjectId }, newReport);
            }
            catch (Exception ex)
            {
                var fullMessage = ex.Message + (ex.InnerException != null ? " | Inner: " + ex.InnerException.Message : "");
                Console.WriteLine($"Error creating envelope lot report: {fullMessage}");
                return StatusCode(500, new { message = "Failed to create report", error = fullMessage });
            }
        }

        // DELETE: api/EnvelopeLotReports/{id}
        [HttpDelete("{id}")]
        public async Task<IActionResult> DeleteEnvelopeLotReport(int id)
        {
            try
            {
                var report = await _context.EnvelopeLotReports.FindAsync(id);
                if (report == null)
                {
                    return NotFound(new { message = "Report not found" });
                }

                _context.EnvelopeLotReports.Remove(report);
                await _context.SaveChangesAsync();

                return Ok(new { message = "Report deleted successfully" });
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { message = "Failed to delete report", error = ex.Message });
            }
        }
    }

    public class CreateEnvelopeLotReportRequest
    {
        [Required]
        public int ProjectId { get; set; }

        [Required]
        public int TemplateId { get; set; }

        [Required]
        public string TemplateName { get; set; }

        // Removed [Required] to allow for project-wide reports or empty selections
        public string EnvLotNumbers { get; set; } = ""; 

        [Required]
        public string FileName { get; set; }

        [Required]
        public string GeneratedBy { get; set; }

        public string? FilePath { get; set; } = null;
    }
}