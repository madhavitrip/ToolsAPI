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
        //[HttpGet("ByProject/{projectId}")]
        //public async Task<ActionResult<IEnumerable<EnvelopeLotReport>>> GetEnvelopeLotReportsByProject(int projectId)
        //{
        //    try
        //    {
        //        Console.WriteLine($"Loading envelope lot reports for project: {projectId}");
        //        var reports = await _context.EnvelopeLotReports
        //            .Where(r => r.ProjectId == projectId)
        //            .OrderByDescending(r => r.GeneratedAt)
        //            .ToListAsync();

        //        Console.WriteLine($"Found {reports.Count} reports for project {projectId}");
        //        return Ok(reports);
        //    }
        //    catch (Exception ex)
        //    {
        //        var fullMessage = ex.Message + (ex.InnerException != null ? " | Inner: " + ex.InnerException.Message : "");
        //        Console.WriteLine($"Error loading reports for project {projectId}: {fullMessage}");
        //        return StatusCode(500, new { message = "Failed to retrieve reports", error = fullMessage });
        //    }
        //}

        [HttpGet("ByProject/{projectId}")]
        public async Task<IActionResult> GetEnvelopeLotReportsByProject(int projectId)
        {
            try
            {
                Console.WriteLine(
                    $"Loading envelope lot reports for project: {projectId}"
                );

                var reports = await _context.EnvelopeLotReports
                    .AsNoTracking()
                    .Where(r => r.ProjectId == projectId)

                    // Join with RPTTemplates using TemplateId
                    .Join(
                        _context.RPTTemplates.AsNoTracking(),
                        envelopeReport => envelopeReport.TemplateId,
                        rptTemplate => rptTemplate.TemplateId,
                        (envelopeReport, rptTemplate) => new
                        {
                            EnvelopeReport = envelopeReport,
                            RPTTemplate = rptTemplate
                        }
                    )

                    .OrderByDescending(x => x.EnvelopeReport.GeneratedAt)
                    .ThenByDescending(x => x.EnvelopeReport.Id)

                    .Select(x => new
                    {
                        x.EnvelopeReport.Id,
                        x.EnvelopeReport.ProjectId,
                        x.EnvelopeReport.TemplateId,
                        x.EnvelopeReport.TemplateName,
                        x.EnvelopeReport.EnvLotNumbers,
                        x.EnvelopeReport.FileName,
                        x.EnvelopeReport.GeneratedAt,
                        x.EnvelopeReport.GeneratedByUserId,
                        x.EnvelopeReport.DownloadedByUserId,
                        x.EnvelopeReport.DownloadedAt,
                        x.EnvelopeReport.FilePath,
                        x.EnvelopeReport.Status,

                        // Version from RPTTemplates table
                        Version = x.RPTTemplate.Version,

                        // Keep this for frontend consistency
                        LotNumber = x.EnvelopeReport.LotNo
                    })
                    .ToListAsync();

                Console.WriteLine(
                    $"Found {reports.Count} EnvelopeLotReports rows " +
                    $"for project {projectId}"
                );

                return Ok(reports);
            }
            catch (Exception ex)
            {
                var fullMessage =
                    ex.Message +
                    (
                        ex.InnerException != null
                            ? " | Inner: " + ex.InnerException.Message
                            : ""
                    );

                Console.WriteLine(
                    $"Error loading reports for project {projectId}: " +
                    fullMessage
                );

                return StatusCode(
                    500,
                    new
                    {
                        message = "Failed to retrieve reports",
                        error = fullMessage
                    }
                );
            }
        }

        //[HttpGet("ByProject/{projectId}")]
        //public async Task<IActionResult> GetEnvelopeLotReportsByProject(int projectId)
        //{
        //    try
        //    {
        //        Console.WriteLine(
        //            $"Loading envelope lot reports for project: {projectId}"
        //        );

        //        // Fetch every EnvelopeLotReports row.
        //        // No GroupBy, Distinct, lookup, or overwrite.
        //        var reports = await _context.EnvelopeLotReports
        //            .AsNoTracking()
        //            .Where(r => r.ProjectId == projectId)
        //            .OrderByDescending(r => r.GeneratedAt)
        //            .ThenByDescending(r => r.Id)
        //            .Select(r => new
        //            {
        //                r.Id,
        //                r.ProjectId,
        //                r.TemplateId,
        //                r.TemplateName,
        //                r.EnvLotNumbers,
        //                r.FileName,
        //                r.GeneratedAt,
        //                r.GeneratedBy,
        //                r.FilePath,
        //                r.LotNo
        //            })
        //            .ToListAsync();

        //        // Get active NR data only for EnvLotDetails.
        //        // This does not affect the report rows themselves.
        //        var nrData = await _context.NRDatas
        //            .AsNoTracking()
        //            .Where(x =>
        //                x.ProjectId == projectId &&
        //                x.Status
        //            )
        //            .Select(x => new
        //            {
        //                x.EnvLotNo,
        //                x.LotNo,
        //                x.CatchNo
        //            })
        //            .ToListAsync();

        //        var result = reports.Select(report =>
        //        {
        //            var envLotNos =
        //                string.IsNullOrWhiteSpace(report.EnvLotNumbers) ||
        //                report.EnvLotNumbers == "0"
        //                    ? new List<int>()
        //                    : report.EnvLotNumbers
        //                        .Split(
        //                            ',',
        //                            StringSplitOptions.RemoveEmptyEntries
        //                        )
        //                        .Select(x =>
        //                        {
        //                            return int.TryParse(
        //                                x.Trim(),
        //                                out var envLotNo
        //                            )
        //                                ? envLotNo
        //                                : 0;
        //                        })
        //                        .Where(x => x > 0)
        //                        .Distinct()
        //                        .ToList();

        //            var envLotDetails = nrData
        //                .Where(n =>
        //                    envLotNos.Contains(n.EnvLotNo)
        //                )
        //                .GroupBy(n => n.EnvLotNo)
        //                .Select(envGroup => new
        //                {
        //                    EnvLotNo = envGroup.Key,

        //                    Lots = envGroup
        //                        .GroupBy(n => n.LotNo)
        //                        .Select(lotGroup => new
        //                        {
        //                            LotNo = lotGroup.Key,

        //                            CatchNos = lotGroup
        //                                .Where(n =>
        //                                    !string.IsNullOrWhiteSpace(
        //                                        n.CatchNo
        //                                    )
        //                                )
        //                                .Select(n => n.CatchNo)
        //                                .Distinct()
        //                                .OrderBy(x => x)
        //                                .ToList()
        //                        })
        //                        .OrderBy(x => x.LotNo)
        //                        .ToList()
        //                })
        //                .OrderBy(x => x.EnvLotNo)
        //                .ToList();

        //            // IMPORTANT:
        //            // Every database row becomes one separate result object.
        //            // LotNo is taken directly from EnvelopeLotReports.
        //            return new
        //            {
        //                report.Id,
        //                report.ProjectId,
        //                report.TemplateId,
        //                report.TemplateName,
        //                report.EnvLotNumbers,
        //                report.LotNo,
        //                report.FileName,
        //                report.GeneratedAt,
        //                report.GeneratedBy,
        //                report.FilePath,

        //                // Optional normalized property for frontend
        //                LotNumber = report.LotNo,

        //                EnvLotDetails = envLotDetails
        //            };
        //        }).ToList();

        //        Console.WriteLine(
        //            $"Found {result.Count} EnvelopeLotReports rows " +
        //            $"for project {projectId}"
        //        );

        //        return Ok(result);
        //    }
        //    catch (Exception ex)
        //    {
        //        var fullMessage =
        //            ex.Message +
        //            (
        //                ex.InnerException != null
        //                    ? " | Inner: " +
        //                      ex.InnerException.Message
        //                    : ""
        //            );

        //        Console.WriteLine(
        //            $"Error loading reports for project {projectId}: " +
        //            fullMessage
        //        );

        //        return StatusCode(
        //            500,
        //            new
        //            {
        //                message = "Failed to retrieve reports",
        //                error = fullMessage
        //            }
        //        );
        //    }
        //}

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
                Console.WriteLine($"Received request to create envelope lot report: ProjectId={request.ProjectId}, TemplateId={request.TemplateId}, EnvLotNumbers={request.EnvLotNumbers}, LotNo={request.LotNo}, GeneratedByUserId={request.GeneratedByUserId}");
                
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
                    LotNo = request.LotNo ?? 0,
                    FileName = request.FileName,
                    GeneratedAt = DateTime.UtcNow,
                    GeneratedByUserId = request.GeneratedByUserId,
                    FilePath = request.FilePath
                };

                _context.EnvelopeLotReports.Add(newReport);
                await _context.SaveChangesAsync();
                Console.WriteLine($"New report created with ID: {newReport.Id}, LotNo: {newReport.LotNo}, GeneratedByUserId: {newReport.GeneratedByUserId}");

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

        // PUT: api/EnvelopeLotReports/{id}/track-download
        [HttpPut("{id}/track-download")]
        public async Task<IActionResult> TrackDownload(int id, [FromBody] DownloadTrackingRequest request)
        {
            try
            {
                var report = await _context.EnvelopeLotReports.FindAsync(id);
                if (report == null)
                {
                    return NotFound(new { message = "Report not found" });
                }

                report.DownloadedByUserId = request.DownloadedByUserId;
                report.DownloadedAt = DateTime.UtcNow;

                _context.EnvelopeLotReports.Update(report);
                await _context.SaveChangesAsync();

                Console.WriteLine($"Tracked download for report {id}: UserId: {report.DownloadedByUserId} at {report.DownloadedAt}");

                return Ok(new 
                { 
                    message = "Download tracked successfully",
                    downloadedByUserId = report.DownloadedByUserId,
                    downloadedAt = report.DownloadedAt
                });
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error tracking download: {ex.Message}");
                return StatusCode(500, new { message = "Failed to track download", error = ex.Message });
            }
        }

        // PUT: api/EnvelopeLotReports/{id}/archive
        [HttpPut("{id}/archive")]
        public async Task<IActionResult> ArchiveReport(int id)
        {
            try
            {
                var report = await _context.EnvelopeLotReports.FindAsync(id);
                if (report == null)
                {
                    return NotFound(new { message = "Report not found" });
                }

                report.Status = false;
                _context.EnvelopeLotReports.Update(report);
                await _context.SaveChangesAsync();

                return Ok(new { message = "Report archived successfully" });
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { message = "Failed to archive report", error = ex.Message });
            }
        }

        // PUT: api/EnvelopeLotReports/{id}/unarchive
        [HttpPut("{id}/unarchive")]
        public async Task<IActionResult> UnarchiveReport(int id)
        {
            try
            {
                var report = await _context.EnvelopeLotReports.FindAsync(id);
                if (report == null)
                {
                    return NotFound(new { message = "Report not found" });
                }

                report.Status = true;
                _context.EnvelopeLotReports.Update(report);
                await _context.SaveChangesAsync();

                return Ok(new { message = "Report unarchived successfully" });
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { message = "Failed to unarchive report", error = ex.Message });
            }
        }
    }

    public class DownloadTrackingRequest
    {
        public int? DownloadedByUserId { get; set; }
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

        public int? LotNo { get; set; } // The actual lot number for lot-based reports

        [Required]
        public string FileName { get; set; }

        public int? GeneratedByUserId { get; set; }

        public string? FilePath { get; set; } = null;
    }
}