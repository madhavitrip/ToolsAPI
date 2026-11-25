using ERPToolsAPI.Data;
using Microsoft.EntityFrameworkCore;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Tools.Models;
using Tools.Services;

namespace ToolsAPI.Helpers
{
    public static class ExcelReportHelper
    {
        public static async Task<ExcelReport?> RecordExcelReportAsync(
            ERPToolsDbContext context,
            int? projectId,
            int moduleId,
            int version,
            int? lot,
            string? filePath,
            bool status,
            int? generatedByUserId = null)
        {
            if (context == null) return null;

            try
            {
                var report = new ExcelReport
                {
                    ProjectId = projectId,
                    ModuleId = moduleId,
                    Version = version,
                    Lot = lot,
                    FilePath = FileStorageHelper.GetRelativePath(filePath),
                    Status = status,
                    GeneratedAt = DateTime.Now,
                    GeneratedByUserId = generatedByUserId > 0 ? generatedByUserId : null
                };

                context.ExcelReports.Add(report);
                await context.SaveChangesAsync();
                return report;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[ExcelReportHelper] Error recording excel report: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Inactivates active ExcelReports for a project and affected lot (or all lots if lotNo is null).
        /// When lotNo is specified, also invalidates project-wide aggregate reports (Lot == null or Lot == 0)
        /// because the aggregate report contains data from the affected lot.
        /// Unaffected lots remain active. Physical files on disk are preserved.
        /// </summary>
        public static async Task InvalidateReportsAsync(
            ERPToolsDbContext context,
            int projectId,
            int? lotNo = null,
            IEnumerable<int>? moduleIds = null)
        {
            if (context == null || projectId <= 0) return;

            try
            {
                var query = context.ExcelReports
                    .Where(r => r.ProjectId == projectId && r.Status);

                if (moduleIds != null && moduleIds.Any())
                {
                    query = query.Where(r => moduleIds.Contains(r.ModuleId));
                }

                if (lotNo.HasValue && lotNo.Value > 0)
                {
                    // Invalidate reports for this specific lot AND project-wide aggregate reports
                    query = query.Where(r => r.Lot == lotNo.Value || r.Lot == null || r.Lot == 0);
                }

                var reportsToDeactivate = await query.ToListAsync();
                if (reportsToDeactivate.Any())
                {
                    foreach (var rep in reportsToDeactivate)
                    {
                        rep.Status = false;
                        context.ExcelReports.Update(rep);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[ExcelReportHelper] Error invalidating reports: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// Inactivates active ExcelReports for multiple affected lots and project-wide aggregate reports.
        /// Unaffected lots remain active. Physical files on disk are preserved.
        /// </summary>
        public static async Task InvalidateReportsForLotsAsync(
            ERPToolsDbContext context,
            int projectId,
            IEnumerable<int> lotNos,
            IEnumerable<int>? moduleIds = null)
        {
            if (context == null || projectId <= 0) return;

            try
            {
                var validLots = lotNos?.Where(l => l > 0).Distinct().ToList() ?? new List<int>();

                var query = context.ExcelReports
                    .Where(r => r.ProjectId == projectId && r.Status);

                if (moduleIds != null && moduleIds.Any())
                {
                    query = query.Where(r => moduleIds.Contains(r.ModuleId));
                }

                if (validLots.Any())
                {
                    query = query.Where(r => (r.Lot.HasValue && validLots.Contains(r.Lot.Value)) || r.Lot == null || r.Lot == 0);
                }

                var reportsToDeactivate = await query.ToListAsync();
                if (reportsToDeactivate.Any())
                {
                    foreach (var rep in reportsToDeactivate)
                    {
                        rep.Status = false;
                        context.ExcelReports.Update(rep);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[ExcelReportHelper] Error invalidating reports for lots: {ex.Message}");
                throw;
            }
        }
    }
}
