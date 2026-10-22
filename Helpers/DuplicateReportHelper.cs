using ERPToolsAPI.Data;
using Microsoft.EntityFrameworkCore;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.Json;
using System.Threading.Tasks;
using Tools.Models;
using Tools.Services;

namespace ToolsAPI.Helpers
{
    public static class DuplicateReportHelper
    {
        public static readonly List<PropertyInfo> BaseProperties = typeof(NRData)
            .GetProperties(BindingFlags.Public | BindingFlags.Instance)
            .Where(p => p.Name != "NRDatas")
            .ToList();

        public static readonly HashSet<string> BasePropertyNames = new HashSet<string>(
            BaseProperties.Select(p => p.Name),
            StringComparer.OrdinalIgnoreCase
        );

        public static Dictionary<string, string> ParseNRDatasExtras(string? json)
        {
            var extras = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            if (string.IsNullOrWhiteSpace(json))
                return extras;

            try
            {
                using var doc = JsonDocument.Parse(json);
                if (doc.RootElement.ValueKind == JsonValueKind.Object)
                {
                    foreach (var prop in doc.RootElement.EnumerateObject())
                    {
                        extras[prop.Name] = prop.Value.ValueKind switch
                        {
                            JsonValueKind.Null or JsonValueKind.Undefined => "",
                            JsonValueKind.String => prop.Value.GetString() ?? "",
                            _ => prop.Value.GetRawText()
                        };
                    }
                }
            }
            catch
            {
                try
                {
                    var dict = Newtonsoft.Json.JsonConvert.DeserializeObject<Dictionary<string, object>>(json);
                    if (dict != null)
                    {
                        foreach (var kvp in dict)
                        {
                            extras[kvp.Key] = kvp.Value?.ToString() ?? "";
                        }
                    }
                }
                catch { }
            }

            return extras;
        }

        public static object? GetNRDataCellValue(NRData item, PropertyInfo prop, Dictionary<string, string> extras)
        {
            var raw = prop.GetValue(item);
            if (raw != null)
            {
                if (prop.PropertyType == typeof(DateTime) || prop.PropertyType == typeof(DateTime?))
                {
                    var dt = (DateTime?)raw;
                    return (dt.HasValue && dt.Value > DateTime.MinValue) ? dt.Value.ToString("yyyy-MM-dd HH:mm:ss") : "";
                }
                if (raw is IEnumerable<int> intList)
                {
                    return string.Join(", ", intList);
                }
                if (raw is string str)
                {
                    if (string.IsNullOrEmpty(str) && extras != null && extras.TryGetValue(prop.Name, out var extVal) && !string.IsNullOrEmpty(extVal))
                    {
                        return extVal;
                    }
                    return str;
                }
                return raw;
            }

            if (extras != null && extras.TryGetValue(prop.Name, out var fallbackVal) && !string.IsNullOrEmpty(fallbackVal))
            {
                return fallbackVal;
            }

            return "";
        }

        public static List<string> CollectExtraHeaderKeys(IEnumerable<NRData> rows)
        {
            var extraHeaderList = new List<string>();
            var extraHeaderSet = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (var row in rows)
            {
                var extras = ParseNRDatasExtras(row.NRDatas);
                foreach (var key in extras.Keys)
                {
                    if (!BasePropertyNames.Contains(key) && extraHeaderSet.Add(key))
                    {
                        extraHeaderList.Add(key);
                    }
                }
            }

            return extraHeaderList;
        }

        public static void PopulateNRDataWorksheet(
            ExcelWorksheet ws,
            List<NRData> rows,
            List<PropertyInfo> baseProperties,
            List<string> extraHeaderList,
            HashSet<int>? deletedRowIds = null,
            HashSet<int>? newRowIds = null)
        {
            int col = 1;
            foreach (var prop in baseProperties)
            {
                ws.Cells[1, col].Value = prop.Name;
                ws.Cells[1, col].Style.Font.Bold = true;
                col++;
            }

            foreach (var key in extraHeaderList)
            {
                ws.Cells[1, col].Value = key;
                ws.Cells[1, col].Style.Font.Bold = true;
                col++;
            }

            int rowIdx = 2;
            foreach (var item in rows)
            {
                var extras = ParseNRDatasExtras(item.NRDatas);
                col = 1;

                foreach (var prop in baseProperties)
                {
                    ws.Cells[rowIdx, col++].Value = GetNRDataCellValue(item, prop, extras);
                }

                foreach (var key in extraHeaderList)
                {
                    extras.TryGetValue(key, out var val);
                    ws.Cells[rowIdx, col++].Value = val ?? "";
                }

                if (deletedRowIds != null && newRowIds != null && col > 1)
                {
                    if (deletedRowIds.Contains(item.Id))
                    {
                        using var range = ws.Cells[rowIdx, 1, rowIdx, col - 1];
                        range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        range.Style.Fill.BackgroundColor.SetColor(Color.LightCoral); // 🔴 old/deleted
                    }
                    else if (newRowIds.Contains(item.Id))
                    {
                        using var range = ws.Cells[rowIdx, 1, rowIdx, col - 1];
                        range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        range.Style.Fill.BackgroundColor.SetColor(Color.LightGreen); // 🟢 new/merged
                    }
                }

                rowIdx++;
            }

            if (ws.Dimension != null)
            {
                ws.Cells[ws.Dimension.Address].AutoFitColumns();
                ws.View.FreezePanes(2, 1);
            }
        }

        public static void GenerateDuplicateReportPackage(
            string filePath,
            List<NRData> reportRows,
            List<NRData> cleanRows,
            HashSet<int>? deletedRowIds = null,
            HashSet<int>? newRowIds = null,
            string reportSheetName = "Merge Report")
        {
            var dir = Path.GetDirectoryName(filePath);
            if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
            {
                Directory.CreateDirectory(dir);
            }

            var extraHeaderList = CollectExtraHeaderKeys(reportRows.Concat(cleanRows));

            using var package = new ExcelPackage();

            // 1. Report Sheet (e.g. Merge Report or Duplicate Report)
            var ws = package.Workbook.Worksheets.Add(reportSheetName);
            PopulateNRDataWorksheet(ws, reportRows, BaseProperties, extraHeaderList, deletedRowIds, newRowIds);

            // 2. Clean Sheet (Final active data)
            var wsClean = package.Workbook.Worksheets.Add("Clean NRData");
            PopulateNRDataWorksheet(wsClean, cleanRows, BaseProperties, extraHeaderList);

            package.SaveAs(new FileInfo(filePath));
        }

        public static void GenerateNewModelDuplicateReportPackage(
            string filePath,
            List<CenterList> reportCenters,
            List<CenterList> cleanCenters,
            Dictionary<int, NrData1> nrDataMap,
            HashSet<int>? deletedCenterIds = null,
            HashSet<int>? newCenterIds = null,
            string reportSheetName = "Merge Report")
        {
            var dir = Path.GetDirectoryName(filePath);
            if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
            {
                Directory.CreateDirectory(dir);
            }

            var extraHeaders = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var c in reportCenters.Concat(cleanCenters))
            {
                if (!string.IsNullOrWhiteSpace(c.CenterData))
                {
                    try
                    {
                        var dict = JsonSerializer.Deserialize<Dictionary<string, string>>(c.CenterData);
                        if (dict != null)
                        {
                            foreach (var key in dict.Keys)
                            {
                                if (!key.Equals("Quantity", StringComparison.OrdinalIgnoreCase) &&
                                    !key.Equals("NRQuantity", StringComparison.OrdinalIgnoreCase))
                                {
                                    extraHeaders.Add(key);
                                }
                            }
                        }
                    }
                    catch { }
                }
            }

            var extraHeaderList = extraHeaders.OrderBy(x => x).ToList();
            var baseHeaders = new List<string>
            {
                "CatchNo", "SubjectName", "CourseName", "NodalCode", "CenterCode",
                "Route", "Quantity", "NRQuantity", "Status"
            };

            ExcelPackage.License.SetNonCommercialPersonal("Tools");
            using var package = new ExcelPackage();

            // 1. Report Sheet
            var ws = package.Workbook.Worksheets.Add(reportSheetName);
            int col = 1;
            foreach (var h in baseHeaders)
            {
                ws.Cells[1, col].Value = h;
                ws.Cells[1, col].Style.Font.Bold = true;
                col++;
            }
            foreach (var key in extraHeaderList)
            {
                ws.Cells[1, col].Value = key;
                ws.Cells[1, col].Style.Font.Bold = true;
                col++;
            }

            int rowIdx = 2;
            foreach (var item in reportCenters)
            {
                nrDataMap.TryGetValue(item.NRDataId, out var catchInfo);

                col = 1;
                ws.Cells[rowIdx, col++].Value = catchInfo?.CatchNo ?? "";
                ws.Cells[rowIdx, col++].Value = catchInfo?.SubjectName ?? "";
                ws.Cells[rowIdx, col++].Value = catchInfo?.CourseName ?? "";
                ws.Cells[rowIdx, col++].Value = item.NodalCode ?? "";
                ws.Cells[rowIdx, col++].Value = item.CenterCode ?? "";
                ws.Cells[rowIdx, col++].Value = item.Route ?? "";
                ws.Cells[rowIdx, col++].Value = item.Quantity;
                ws.Cells[rowIdx, col++].Value = item.NRQuantity;
                ws.Cells[rowIdx, col++].Value = item.Status ? "Active" : "Inactive";

                Dictionary<string, string>? cDict = null;
                if (!string.IsNullOrWhiteSpace(item.CenterData))
                {
                    try { cDict = JsonSerializer.Deserialize<Dictionary<string, string>>(item.CenterData); } catch { }
                }

                foreach (var key in extraHeaderList)
                {
                    string val = "";
                    if (cDict != null && cDict.TryGetValue(key, out var dynVal))
                    {
                        val = dynVal ?? "";
                    }
                    ws.Cells[rowIdx, col++].Value = val;
                }

                if (deletedCenterIds != null && deletedCenterIds.Contains(item.Id))
                {
                    using var range = ws.Cells[rowIdx, 1, rowIdx, col - 1];
                    range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    range.Style.Fill.BackgroundColor.SetColor(Color.LightCoral);
                }
                else if (newCenterIds != null && newCenterIds.Contains(item.Id))
                {
                    using var range = ws.Cells[rowIdx, 1, rowIdx, col - 1];
                    range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    range.Style.Fill.BackgroundColor.SetColor(Color.LightGreen);
                }

                rowIdx++;
            }

            if (ws.Dimension != null)
            {
                ws.Cells[ws.Dimension.Address].AutoFitColumns();
                ws.View.FreezePanes(2, 1);
            }

            // 2. Clean Sheet
            var wsClean = package.Workbook.Worksheets.Add("Clean Data");
            col = 1;
            foreach (var h in baseHeaders)
            {
                wsClean.Cells[1, col].Value = h;
                wsClean.Cells[1, col].Style.Font.Bold = true;
                col++;
            }
            foreach (var key in extraHeaderList)
            {
                wsClean.Cells[1, col].Value = key;
                wsClean.Cells[1, col].Style.Font.Bold = true;
                col++;
            }

            int cleanRow = 2;
            foreach (var item in cleanCenters)
            {
                nrDataMap.TryGetValue(item.NRDataId, out var catchInfo);

                col = 1;
                wsClean.Cells[cleanRow, col++].Value = catchInfo?.CatchNo ?? "";
                wsClean.Cells[cleanRow, col++].Value = catchInfo?.SubjectName ?? "";
                wsClean.Cells[cleanRow, col++].Value = catchInfo?.CourseName ?? "";
                wsClean.Cells[cleanRow, col++].Value = item.NodalCode ?? "";
                wsClean.Cells[cleanRow, col++].Value = item.CenterCode ?? "";
                wsClean.Cells[cleanRow, col++].Value = item.Route ?? "";
                wsClean.Cells[cleanRow, col++].Value = item.Quantity;
                wsClean.Cells[cleanRow, col++].Value = item.NRQuantity;
                wsClean.Cells[cleanRow, col++].Value = "Active";

                Dictionary<string, string>? cDict = null;
                if (!string.IsNullOrWhiteSpace(item.CenterData))
                {
                    try { cDict = JsonSerializer.Deserialize<Dictionary<string, string>>(item.CenterData); } catch { }
                }

                foreach (var key in extraHeaderList)
                {
                    string val = "";
                    if (cDict != null && cDict.TryGetValue(key, out var dynVal))
                    {
                        val = dynVal ?? "";
                    }
                    wsClean.Cells[cleanRow, col++].Value = val;
                }

                cleanRow++;
            }

            if (wsClean.Dimension != null)
            {
                wsClean.Cells[wsClean.Dimension.Address].AutoFitColumns();
                wsClean.View.FreezePanes(2, 1);
            }

            package.SaveAs(new FileInfo(filePath));
        }

        public static async Task<string> RegenerateReportFileAsync(
            ERPToolsDbContext context,
            int projectId,
            int? lotNo,
            string targetPath)
        {
            var fullPath = FileStorageHelper.GetAbsolutePath(targetPath);
            var dir = Path.GetDirectoryName(fullPath);
            if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
            {
                Directory.CreateDirectory(dir);
            }

            var isNewModel = await context.NrData1.AnyAsync(p => p.ProjectId == projectId && p.Status);
            if (isNewModel)
            {
                var nrDataMap = await context.NrData1
                    .Where(d => d.ProjectId == projectId && d.Status)
                    .ToDictionaryAsync(d => d.Id);

                var cleanCentersQuery = context.CenterList
                    .Where(c => c.ProjectId == projectId && c.Status && c.ExtraId == 0);

                if (lotNo.HasValue && lotNo.Value > 0)
                {
                    var matchingCatchIds = nrDataMap.Values
                        .Where(nr => nr.LotNo == lotNo.Value || nr.EnvLotNo == lotNo.Value)
                        .Select(nr => nr.Id)
                        .ToHashSet();
                    cleanCentersQuery = cleanCentersQuery.Where(c => matchingCatchIds.Contains(c.NRDataId));
                }

                var cleanCenters = await cleanCentersQuery.ToListAsync();

                var allCentersQuery = context.CenterList
                    .Where(c => c.ProjectId == projectId && c.ExtraId == 0);
                if (lotNo.HasValue && lotNo.Value > 0)
                {
                    var matchingCatchIds = nrDataMap.Values
                        .Where(nr => nr.LotNo == lotNo.Value || nr.EnvLotNo == lotNo.Value)
                        .Select(nr => nr.Id)
                        .ToHashSet();
                    allCentersQuery = allCentersQuery.Where(c => matchingCatchIds.Contains(c.NRDataId));
                }
                var allCenters = await allCentersQuery.ToListAsync();

                var deletedCenters = allCenters.Where(c => !c.Status).ToList();
                var reportCenters = deletedCenters.Any() ? allCenters : cleanCenters;

                GenerateNewModelDuplicateReportPackage(
                    fullPath,
                    reportCenters,
                    cleanCenters,
                    nrDataMap,
                    deletedCenters.Select(c => c.Id).ToHashSet(),
                    null,
                    "Merge Report"
                );

                return fullPath;
            }

            // Legacy fallback for old projects
            var cleanRowsQuery = context.NRDatas
                .Where(x => x.ProjectId == projectId && x.Status == true);
            if (lotNo.HasValue && lotNo.Value > 0)
            {
                cleanRowsQuery = cleanRowsQuery.Where(x => x.LotNo == lotNo.Value);
            }
            var cleanRows = await cleanRowsQuery.ToListAsync();

            // All records for the lot (including merged or deactivated rows) to build Merge Report
            var allLotRowsQuery = context.NRDatas
                .Where(x => x.ProjectId == projectId);
            if (lotNo.HasValue && lotNo.Value > 0)
            {
                allLotRowsQuery = allLotRowsQuery.Where(x => x.LotNo == lotNo.Value);
            }
            var allLotRows = await allLotRowsQuery.ToListAsync();

            var deletedRowIds = new HashSet<int>(allLotRows.Where(x => !x.Status).Select(x => x.Id));
            var newRowIds = new HashSet<int>(allLotRows.Where(x => x.Status && x.Steps == PipelineNavigator.STEP_DUP_PARTIAL).Select(x => x.Id));

            // If there are duplicate records, reportRows is allLotRows; otherwise cleanRows
            var reportRows = deletedRowIds.Any() || newRowIds.Any() ? allLotRows : cleanRows;

            GenerateDuplicateReportPackage(
                fullPath,
                reportRows,
                cleanRows,
                deletedRowIds,
                newRowIds,
                "Merge Report"
            );

            return fullPath;
        }

        public static bool ShouldRegenerateDuplicateReport(string filePath)
        {
            if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath))
                return true;

            try
            {
                using var package = new ExcelPackage(new FileInfo(filePath));
                var cleanWs = package.Workbook.Worksheets["Clean Data"] ?? package.Workbook.Worksheets["Clean NRData"];
                if (cleanWs == null || cleanWs.Dimension == null)
                    return true;

                // If at least 1 data row exists, do not trigger regeneration
                if (cleanWs.Dimension.End.Row < 2)
                    return true;

                return false;
            }
            catch
            {
                return true;
            }
        }
    }
}
