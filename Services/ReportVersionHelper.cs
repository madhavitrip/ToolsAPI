using System;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using ERPToolsAPI.Data;
using Microsoft.EntityFrameworkCore;

namespace Tools.Services
{
    public static class ReportVersionHelper
    {
        public static string GetNextVersionFileName(string directory, string baseFileName)
        {
            return GetNextVersionFileName(directory, baseFileName, 0);
        }

        public static string GetNextVersionFileName(string directory, string baseFileName, int minExistingVersion)
        {
            if (string.IsNullOrWhiteSpace(baseFileName))
                return baseFileName;

            string extension = Path.GetExtension(baseFileName);
            string nameWithoutExt = Path.GetFileNameWithoutExtension(baseFileName);

            int maxVersion = Math.Max(0, minExistingVersion);

            if (Directory.Exists(directory))
            {
                // Pattern matches nameWithoutExt_v{number}.extension (e.g. DuplicateTool_v2.xlsx)
                var pattern = new Regex(@"^" + Regex.Escape(nameWithoutExt) + @"_v(\d+)" + Regex.Escape(extension) + "$", RegexOptions.IgnoreCase);
                foreach (var file in Directory.GetFiles(directory))
                {
                    var match = pattern.Match(Path.GetFileName(file));
                    if (match.Success)
                    {
                        if (int.TryParse(match.Groups[1].Value, out int v))
                        {
                            if (v > maxVersion)
                            {
                                maxVersion = v;
                            }
                        }
                    }
                }
            }

            // If maxVersion is 0, check if the unversioned file exists
            if (maxVersion == 0)
            {
                var unversionedPath = Path.Combine(directory, baseFileName);
                if (File.Exists(unversionedPath))
                {
                    try
                    {
                        var v1Path = Path.Combine(directory, $"{nameWithoutExt}_v1{extension}");
                        if (!File.Exists(v1Path))
                        {
                            File.Move(unversionedPath, v1Path);
                        }
                        maxVersion = 1;
                    }
                    catch
                    {
                        maxVersion = 0;
                    }
                }
            }

            int nextVersion = maxVersion + 1;
            return $"{nameWithoutExt}_v{nextVersion}{extension}";
        }

        /// <summary>
        /// Computes the maximum registered version in ExcelReports (both active and inactive records)
        /// for a specific project and report base filename series (e.g. "EnvelopeBreaking_New_5.xlsx").
        /// </summary>
        public static async Task<int> GetDatabaseMaxVersionAsync(
            ERPToolsDbContext context,
            int projectId,
            string baseFileName)
        {
            if (context == null || projectId <= 0 || string.IsNullOrWhiteSpace(baseFileName))
                return 0;

            string extension = Path.GetExtension(baseFileName);
            string nameWithoutExt = Path.GetFileNameWithoutExtension(baseFileName);
            var pattern = new Regex(
                @"^" + Regex.Escape(nameWithoutExt) + @"_v(\d+)" + Regex.Escape(extension) + "$",
                RegexOptions.IgnoreCase);

            // Include both active (Status == true) and inactive (Status == false) historical records
            var records = await context.ExcelReports
                .AsNoTracking()
                .Where(r => r.ProjectId == projectId && r.FilePath != null)
                .Select(r => new { r.FilePath, r.Version })
                .ToListAsync();

            int dbMax = 0;
            foreach (var rec in records)
            {
                var fileName = Path.GetFileName(rec.FilePath!);
                if (string.IsNullOrEmpty(fileName))
                    continue;

                var match = pattern.Match(fileName);
                if (match.Success && int.TryParse(match.Groups[1].Value, out int fileVer))
                {
                    int effectiveVer = Math.Max(fileVer, rec.Version);
                    if (effectiveVer > dbMax)
                        dbMax = effectiveVer;
                }
                else if (string.Equals(fileName, baseFileName, StringComparison.OrdinalIgnoreCase))
                {
                    int effectiveVer = Math.Max(1, rec.Version);
                    if (effectiveVer > dbMax)
                        dbMax = effectiveVer;
                }
            }

            return dbMax;
        }

        /// <summary>
        /// Calculates NextVersion = MAX(DatabaseMaxVersion, PhysicalFileMaxVersion) + 1
        /// for the given project, storage directory, and report base filename series.
        /// </summary>
        public static async Task<(string FileName, int Version)> GetNextVersionAsync(
            ERPToolsDbContext context,
            int projectId,
            string directory,
            string baseFileName)
        {
            int dbMax = await GetDatabaseMaxVersionAsync(context, projectId, baseFileName);
            string nextFileName = GetNextVersionFileName(directory, baseFileName, dbMax);
            int nextVersion = ParseVersionFromFileName(nextFileName, dbMax + 1);
            return (nextFileName, nextVersion);
        }

        public static int ParseVersionFromFileName(string? fileName, int fallback = 1)
        {
            if (string.IsNullOrWhiteSpace(fileName))
                return fallback;

            var match = Regex.Match(fileName, @"_v(\d+)\.[^.]+$", RegexOptions.IgnoreCase);
            if (match.Success && int.TryParse(match.Groups[1].Value, out int parsed) && parsed > 0)
                return parsed;

            return fallback;
        }
    }
}
