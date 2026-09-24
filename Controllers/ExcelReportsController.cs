using ERPToolsAPI.Data;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
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

            // Fallback: If no reports in table or to discover unrecorded excel files directly from disk
            var diskReports = FallbackScanExcelFiles(projectId);
            if (diskReports.Any())
            {
                var existingPaths = new HashSet<string>(
                    reports.Select(r => FileStorageHelper.GetRelativePath(r.FilePath ?? "")),
                    StringComparer.OrdinalIgnoreCase
                );

                foreach (var diskReport in diskReports)
                {
                    if (!existingPaths.Contains(diskReport.FilePath))
                    {
                        reports.Add(diskReport);
                    }
                }
            }

            bool changedInDb = false;
            foreach (var report in reports)
            {
                var fileName = (report.FilePath ?? "").ToLowerInvariant();
                if (fileName.Contains("enhancement") && report.ModuleId != 2)
                {
                    report.ModuleId = 2;
                    if (report.Id > 0)
                    {
                        _context.Entry(report).Property(r => r.ModuleId).IsModified = true;
                        changedInDb = true;
                    }
                }
                else if (fileName.Contains("extra") && report.ModuleId != 3)
                {
                    report.ModuleId = 3;
                    if (report.Id > 0)
                    {
                        _context.Entry(report).Property(r => r.ModuleId).IsModified = true;
                        changedInDb = true;
                    }
                }
            }

            if (changedInDb)
            {
                try
                {
                    await _context.SaveChangesAsync();
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"[ExcelReportsController] Error saving normalized ModuleIds: {ex.Message}");
                }
            }

            return Ok(reports.OrderByDescending(r => r.GeneratedAt));
        }

        // GET: api/ExcelReports/ByModule/{moduleId}/{projectId}
        [HttpGet("ByModule/{moduleId}/{projectId}")]
        public async Task<ActionResult<IEnumerable<ExcelReport>>> GetReportsByModule(int moduleId, int projectId)
        {
            var reports = await _context.ExcelReports
                .Where(r => r.ProjectId == projectId)
                .OrderByDescending(r => r.GeneratedAt)
                .ToListAsync();

            // Fallback: Fetch excel directly from storage directory for this project
            var diskReports = FallbackScanExcelFiles(projectId);
            if (diskReports.Any())
            {
                var existingPaths = new HashSet<string>(
                    reports.Select(r => FileStorageHelper.GetRelativePath(r.FilePath ?? "")),
                    StringComparer.OrdinalIgnoreCase
                );

                foreach (var diskReport in diskReports)
                {
                    if (!existingPaths.Contains(diskReport.FilePath))
                    {
                        reports.Add(diskReport);
                    }
                }
            }

            bool changedInDb = false;
            foreach (var report in reports)
            {
                var fileName = (report.FilePath ?? "").ToLowerInvariant();
                if (fileName.Contains("enhancement") && report.ModuleId != 2)
                {
                    report.ModuleId = 2;
                    if (report.Id > 0)
                    {
                        _context.Entry(report).Property(r => r.ModuleId).IsModified = true;
                        changedInDb = true;
                    }
                }
                else if (fileName.Contains("extra") && report.ModuleId != 3)
                {
                    report.ModuleId = 3;
                    if (report.Id > 0)
                    {
                        _context.Entry(report).Property(r => r.ModuleId).IsModified = true;
                        changedInDb = true;
                    }
                }
            }

            if (changedInDb)
            {
                try
                {
                    await _context.SaveChangesAsync();
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"[ExcelReportsController] Error saving normalized ModuleIds: {ex.Message}");
                }
            }

            var filteredReports = reports.Where(r => r.ModuleId == moduleId).OrderByDescending(r => r.GeneratedAt);
            return Ok(filteredReports);
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
            ExcelReport? report = null;
            string? fullPath = null;

            if (id > 0)
            {
                report = await _context.ExcelReports.FindAsync(id);
                if (report != null && !string.IsNullOrWhiteSpace(report.FilePath))
                {
                    fullPath = FileStorageHelper.GetAbsolutePath(report.FilePath);
                }
            }

            // Fallback if record not found in DB or file doesn't exist at DB path
            if (string.IsNullOrEmpty(fullPath) || !System.IO.File.Exists(fullPath))
            {
                fullPath = FindFallbackExcelFileByIdOrPath(id, report?.FilePath, report?.ProjectId);
            }

            if (string.IsNullOrEmpty(fullPath) || !System.IO.File.Exists(fullPath))
            {
                return NotFound($"Report file not found on server.");
            }

            var mimeType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            var fileNameToReturn = System.IO.Path.GetFileName(fullPath);
            var fileStream = System.IO.File.OpenRead(fullPath);
            return File(fileStream, mimeType, fileNameToReturn);
        }

        // GET: api/ExcelReports/DownloadByPath
        [HttpGet("DownloadByPath")]
        public IActionResult DownloadReportByPath([FromQuery] string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath))
                return BadRequest("filePath parameter is required.");

            var fullPath = FileStorageHelper.GetAbsolutePath(filePath);
            if (string.IsNullOrEmpty(fullPath) || !System.IO.File.Exists(fullPath))
            {
                fullPath = FileStorageHelper.FindExistingFilePath(filePath);
            }

            if (string.IsNullOrEmpty(fullPath) || !System.IO.File.Exists(fullPath))
            {
                return NotFound("Excel report file not found on server.");
            }

            var mimeType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            var fileNameToReturn = System.IO.Path.GetFileName(fullPath);
            var fileStream = System.IO.File.OpenRead(fullPath);
            return File(fileStream, mimeType, fileNameToReturn);
        }

        #region Helper Methods for Fallback Disk Scanning

        private List<ExcelReport> FallbackScanExcelFiles(int projectId, int? targetModuleId = null)
        {
            var reports = new List<ExcelReport>();
            var folderPath = FileStorageHelper.GetProjectFolder(projectId);
            if (!Directory.Exists(folderPath))
            {
                return reports;
            }

            try
            {
                var files = Directory.GetFiles(folderPath, "*.*", SearchOption.AllDirectories)
                    .Where(f => f.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) || f.EndsWith(".xls", StringComparison.OrdinalIgnoreCase))
                    .ToList();

                int index = 1;
                foreach (var file in files)
                {
                    var fileInfo = new FileInfo(file);
                    var fileName = Path.GetFileName(file);
                    var relativePath = FileStorageHelper.GetRelativePath(file);

                    int detectedModuleId = DetectModuleIdFromFileName(fileName);
                    if (targetModuleId.HasValue && targetModuleId.Value > 0 && detectedModuleId > 0 && detectedModuleId != targetModuleId.Value)
                    {
                        continue;
                    }

                    int version = 1;
                    var versionMatch = Regex.Match(fileName, @"_v(\d+)", RegexOptions.IgnoreCase);
                    if (versionMatch.Success && int.TryParse(versionMatch.Groups[1].Value, out int parsedVer))
                    {
                        version = parsedVer;
                    }

                    int? lot = null;
                    var lotMatch = Regex.Match(fileName, @"_(\d+)(?:_v\d+)?\.(?:xlsx|xls)$", RegexOptions.IgnoreCase);
                    if (lotMatch.Success && int.TryParse(lotMatch.Groups[1].Value, out int parsedLot))
                    {
                        lot = parsedLot;
                    }

                    int syntheticId = -Math.Abs(relativePath.GetHashCode());
                    if (syntheticId == 0) syntheticId = -index;

                    reports.Add(new ExcelReport
                    {
                        Id = syntheticId,
                        ProjectId = projectId,
                        ModuleId = detectedModuleId > 0 ? detectedModuleId : (targetModuleId ?? 0),
                        Version = version,
                        Lot = lot,
                        FilePath = relativePath,
                        Status = true,
                        GeneratedAt = fileInfo.LastWriteTime,
                        GeneratedByUserId = null
                    });

                    index++;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[ExcelReportsController] Fallback scanning error: {ex.Message}");
            }

            return reports;
        }

        private int DetectModuleIdFromFileName(string fileName)
        {
            var fn = fileName.ToLowerInvariant();
            if (fn.Contains("duplicate")) return 1;
            if (fn.Contains("enhancement")) return 2;
            if (fn.Contains("extra")) return 3;
            if (fn.Contains("envelopebreaking") || fn.Contains("envelope_breaking")) return 4;
            if (fn.Contains("boxbreaking") || fn.Contains("box_breaking")) return 5;
            if (fn.Contains("envelopesummary") || fn.Contains("envelope_summary")) return 6;
            if (fn.Contains("catchsummary") || fn.Contains("catch_summary") || fn.Contains("catchwise")) return 7;
            return 0;
        }

        private string? FindFallbackExcelFileByIdOrPath(int id, string? filePath, int? projectId)
        {
            if (!string.IsNullOrWhiteSpace(filePath))
            {
                var path = FileStorageHelper.GetAbsolutePath(filePath);
                if (!string.IsNullOrEmpty(path) && System.IO.File.Exists(path))
                    return path;

                var existing = FileStorageHelper.FindExistingFilePath(filePath);
                if (existing != null && System.IO.File.Exists(existing))
                    return existing;
            }

            var basePath = FileStorageHelper.GetStorageBasePath();
            if (projectId.HasValue && projectId.Value > 0)
            {
                var projFolder = FileStorageHelper.GetProjectFolder(projectId.Value);
                if (Directory.Exists(projFolder))
                {
                    var projFiles = Directory.GetFiles(projFolder, "*.*", SearchOption.AllDirectories)
                        .Where(f => f.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) || f.EndsWith(".xls", StringComparison.OrdinalIgnoreCase));

                    foreach (var file in projFiles)
                    {
                        var relPath = FileStorageHelper.GetRelativePath(file);
                        int hashId = -Math.Abs(relPath.GetHashCode());
                        if (hashId == id)
                        {
                            return file;
                        }
                    }
                }
            }

            if (Directory.Exists(basePath))
            {
                var allFiles = Directory.GetFiles(basePath, "*.*", SearchOption.AllDirectories)
                    .Where(f => f.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) || f.EndsWith(".xls", StringComparison.OrdinalIgnoreCase));

                foreach (var file in allFiles)
                {
                    var relPath = FileStorageHelper.GetRelativePath(file);
                    int hashId = -Math.Abs(relPath.GetHashCode());
                    if (hashId == id)
                    {
                        return file;
                    }
                }
            }

            return null;
        }

        #endregion
    }
}
