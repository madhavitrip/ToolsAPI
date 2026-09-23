using ERPToolsAPI.Data;
using System;
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
    }
}
