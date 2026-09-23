using ERPToolsAPI.Data;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Tools.Models;
using Tools.Middleware;
using Tools.Services;

namespace Tools.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ExcelReportsController : ControllerBase
    {
        private readonly ERPToolsDbContext _context;

        public ExcelReportsController(ERPToolsDbContext context)
        {
            _context = context;
        }

        // GET: api/ExcelReports/ByProject/{projectId}
        [HttpGet("ByProject/{projectId}")]
        public async Task<ActionResult<IEnumerable<ExcelReport>>> GetReportsByProject(int projectId)
        {
            var reports = await _context.ExcelReports
                .Where(r => r.ProjectId == projectId)
                .OrderByDescending(r => r.GeneratedAt)
                .ToListAsync();

            return Ok(reports);
        }

        // GET: api/ExcelReports/ByModule/{moduleId}/{projectId}
        [HttpGet("ByModule/{moduleId}/{projectId}")]
        public async Task<ActionResult<IEnumerable<ExcelReport>>> GetReportsByModule(int moduleId, int projectId)
        {
            var reports = await _context.ExcelReports
                .Where(r => r.ProjectId == projectId && r.ModuleId == moduleId)
                .OrderByDescending(r => r.GeneratedAt)
                .ToListAsync();

            return Ok(reports);
        }

        // POST: api/ExcelReports
        [HttpPost]
        public async Task<ActionResult<ExcelReport>> CreateReport([FromBody] ExcelReport report)
        {
            if (report == null) return BadRequest("Invalid report data.");

            report.GeneratedAt = DateTime.Now;
            if (!report.GeneratedByUserId.HasValue)
            {
                report.GeneratedByUserId = LogHelper.GetTriggeredBy(User);
            }

            if (!string.IsNullOrWhiteSpace(report.FilePath))
            {
                report.FilePath = FileStorageHelper.GetRelativePath(report.FilePath);
            }

            _context.ExcelReports.Add(report);
            await _context.SaveChangesAsync();

            return CreatedAtAction(nameof(GetReportsByProject), new { projectId = report.ProjectId }, report);
        }

        // DELETE: api/ExcelReports/{id}
        [HttpDelete("{id}")]
        public async Task<IActionResult> DeleteReport(int id)
        {
            var report = await _context.ExcelReports.FindAsync(id);
            if (report == null)
                return NotFound("Excel report record not found.");

            _context.ExcelReports.Remove(report);
            await _context.SaveChangesAsync();

            return Ok(new { message = "Excel report record deleted successfully." });
        }

        // GET: api/ExcelReports/Download/{id}
        [HttpGet("Download/{id}")]
        public async Task<IActionResult> DownloadReport(int id)
        {
            var report = await _context.ExcelReports.FindAsync(id);
            if (report == null || string.IsNullOrWhiteSpace(report.FilePath))
                return NotFound("Excel report record or file path not found.");

            var fullPath = FileStorageHelper.GetAbsolutePath(report.FilePath);
            if (string.IsNullOrEmpty(fullPath) || !System.IO.File.Exists(fullPath))
            {
                var fileName = System.IO.Path.GetFileName(report.FilePath);
                var relativePath = report.ProjectId.HasValue ? System.IO.Path.Combine(report.ProjectId.Value.ToString(), fileName) : fileName;
                var fallbackPath = FileStorageHelper.FindExistingFilePath(relativePath);
                if (fallbackPath != null && System.IO.File.Exists(fallbackPath))
                {
                    fullPath = fallbackPath;
                }
                else
                {
                    return NotFound($"Report file not found on server.");
                }
            }

            var mimeType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            var fileNameToReturn = System.IO.Path.GetFileName(fullPath);
            var fileStream = System.IO.File.OpenRead(fullPath);
            return File(fileStream, mimeType, fileNameToReturn);
        }
    }
}
