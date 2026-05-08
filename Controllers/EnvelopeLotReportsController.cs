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

       

        // GET: api/EnvelopeLotReports/ByProject/{projectId}
        [HttpGet("ByProject/{projectId}")]
        public async Task<ActionResult<IEnumerable<EnvelopeLotReport>>> GetEnvelopeLotReportsByProject(int projectId)
        {
            try
            {
                Console.WriteLine($"Loading envelope lot reports for project: {projectId}");
                var reports = await _context.EnvelopeLotReports
                    .Where(r => r.ProjectId == projectId)
                    .Select(r => new EnvelopeLotReport
                    {
                        Id = r.Id,
                        ProjectId = r.ProjectId,
                        TemplateId = r.TemplateId,
                        TemplateName = r.TemplateName ?? "",
                        EnvLotNumbers = r.EnvLotNumbers ?? "",
                        FileName = r.FileName ?? "",
                        GeneratedAt = r.GeneratedAt,
                        GeneratedBy = r.GeneratedBy ?? "",
                        FilePath = r.FilePath // This can be null
                    })
                    .OrderByDescending(r => r.GeneratedAt)
                    .ToListAsync();

                Console.WriteLine($"Found {reports.Count} reports for project {projectId}");
                return Ok(reports);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error loading reports for project {projectId}: {ex.Message}");
                Console.WriteLine($"Stack trace: {ex.StackTrace}");
                return StatusCode(500, new { message = "Failed to retrieve reports", error = ex.Message });
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

                // Check if a report with the same template and envelope lots already exists
                var existingReport = await _context.EnvelopeLotReports
                    .FirstOrDefaultAsync(r => r.ProjectId == request.ProjectId 
                                           && r.TemplateId == request.TemplateId 
                                           && r.EnvLotNumbers == request.EnvLotNumbers);

                if (existingReport != null)
                {
                    Console.WriteLine($"Updating existing report with ID: {existingReport.Id}");
                    // Update existing report
                    existingReport.FileName = request.FileName;
                    existingReport.GeneratedAt = DateTime.Now;
                    existingReport.GeneratedBy = request.GeneratedBy;
                    existingReport.FilePath = request.FilePath;

                    await _context.SaveChangesAsync();
                    Console.WriteLine("Report updated successfully");
                    return Ok(existingReport);
                }
                else
                {
                    Console.WriteLine("Creating new report");
                    // Create new report
                    var newReport = new EnvelopeLotReport
                    {
                        ProjectId = request.ProjectId,
                        TemplateId = request.TemplateId,
                        TemplateName = request.TemplateName,
                        EnvLotNumbers = request.EnvLotNumbers,
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
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error creating envelope lot report: {ex.Message}");
                Console.WriteLine($"Stack trace: {ex.StackTrace}");
                return StatusCode(500, new { message = "Failed to create report", error = ex.Message });
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

        [Required]
        public string EnvLotNumbers { get; set; } // Comma-separated envelope lot numbers

        [Required]
        public string FileName { get; set; }

        [Required]
        public string GeneratedBy { get; set; }

        public string? FilePath { get; set; } = null; // Make explicitly nullable with default
    }
}