using ERPToolsAPI.Data;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using System.Reflection;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;
using Tools.Models;
using System.Text.Json;
using Humanizer;
using Newtonsoft.Json;
using NRData = Tools.Models.NRData;
using Tools.Services;

namespace Tools.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class DuplicateController : ControllerBase
    {
        private readonly ERPToolsDbContext _context;
        private readonly ILoggerService _logger;

        public DuplicateController(ERPToolsDbContext context, ILoggerService loggerService)
        {
            _context = context;
            _logger = loggerService;
        }

        [HttpPost]
        public async Task<IActionResult> MergeFields(int ProjectId)
        {
            try
            {
                var data = await _context.NRDatas
                    .Where(p => p.ProjectId == ProjectId && p.Status == true && p.Steps==0)
                    .ToListAsync();
                var projectconfig = await _context.ProjectConfigs
                    .FirstOrDefaultAsync(p => p.ProjectId == ProjectId);
                if (projectconfig == null)
                    return NotFound("Project config not exists for this project");

                var mergeFieldIds = projectconfig.DuplicateCriteria ?? new List<int>();
                if (!mergeFieldIds.Any())
                    return BadRequest("Duplicate criteria is not configured for this project.");

                var fieldNames = await _context.Fields
                    .Where(f => mergeFieldIds.Contains(f.FieldId))
                    .Select(f => f.Name)
                    .ToListAsync();
                if (!fieldNames.Any())
                    return BadRequest("Duplicate criteria fields not found.");

                var grouped = data.GroupBy(d =>
                {
                    var key = new List<string>();

                    foreach (var field in fieldNames)
                    {
                        var value = d.GetType()
                            .GetProperty(field, BindingFlags.IgnoreCase | BindingFlags.Public | BindingFlags.Instance)
                            ?.GetValue(d)?.ToString()?.Trim() ?? "";

                        key.Add(value);
                    }

                    return string.Join("|", key);
                });

                int mergedCount = 0;
                var deletedRows = new List<NRData>();
                var reportRows = new List<NRData>();
                var newRows = new List<NRData>(); // ✅ track new rows

                foreach (var group in grouped)
                {
                    reportRows.AddRange(group);

                    if (group.Count() <= 1)
                        continue;

                    var first = group.First();

                    // 🔹 Clone
                    var newRow = new NRData();

                    foreach (var prop in typeof(NRData).GetProperties(BindingFlags.Public | BindingFlags.Instance))
                    {
                        if (!prop.CanWrite || prop.Name == "Id")
                            continue;

                        prop.SetValue(newRow, prop.GetValue(first));
                    }

                    // 🔹 Override merged values
                    newRow.Status = true;
                    newRow.NRQuantity = group.Sum(x => x.NRQuantity);
                    newRow.Steps = 1;
                    var subjectValues = group.Select(x => x.SubjectName?.Trim())
                        .Where(v => !string.IsNullOrEmpty(v))
                        .Distinct(StringComparer.OrdinalIgnoreCase)
                        .ToList();

                    newRow.SubjectName = subjectValues.Count > 1
                        ? string.Join(" / ", subjectValues)
                        : subjectValues.FirstOrDefault();

                    var courseValues = group.Select(x => x.CourseName?.Trim())
                        .Where(v => !string.IsNullOrEmpty(v))
                        .Distinct(StringComparer.OrdinalIgnoreCase)
                        .ToList();

                    newRow.CourseName = courseValues.Count > 1
                        ? string.Join(" / ", courseValues)
                        : courseValues.FirstOrDefault();

                    await _context.NRDatas.AddAsync(newRow);
                    newRows.Add(newRow); // ✅ track

                    foreach (var item in group)
                    {
                        item.Status = false;
                        deletedRows.Add(item);
                    }

                    mergedCount += group.Count() - 1;
                }
                foreach (var row in data)
                {
                    row.Steps = 1;
                }
                await _context.SaveChangesAsync();

                // ✅ Add new rows to report AFTER save (IDs generated)
                reportRows.AddRange(newRows);

                var triggeredBy = LogHelper.GetTriggeredBy(User);

                await _logger.LogEventAsync(
                    "Duplicates merged and new rows created",
                    "Duplicates",
                    triggeredBy,
                    ProjectId,
                    string.Empty,
                    LogHelper.ToJson(new { ProjectId, MergedCount = mergedCount })
                );

                // ================= REPORT =================

                var reportPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", ProjectId.ToString());
                if (!Directory.Exists(reportPath))
                    Directory.CreateDirectory(reportPath);

                var filePath = Path.Combine(reportPath, "DuplicateTool.xlsx");
                if (System.IO.File.Exists(filePath))
                    System.IO.File.Delete(filePath);

                var baseProperties = typeof(NRData).GetProperties(BindingFlags.Public | BindingFlags.Instance)
                    .Where(p => p.Name != "NRDatas" && p.Name != "Id" && p.Name != "ProjectId")
                    .Where(p => reportRows.Any(row => p.GetValue(row) != null && !string.IsNullOrEmpty(p.GetValue(row)?.ToString())))
                    .ToList();

                var extraHeaders = new HashSet<string>();
                var parsedRows = new List<(NRData row, Dictionary<string, string> extras)>();

                foreach (var row in reportRows)
                {
                    var extras = new Dictionary<string, string>();

                    if (!string.IsNullOrEmpty(row.NRDatas))
                    {
                        try
                        {
                            extras = System.Text.Json.JsonSerializer.Deserialize<Dictionary<string, string>>(row.NRDatas);
                            if (extras != null)
                            {
                                foreach (var key in extras.Keys)
                                    extraHeaders.Add(key);
                            }
                        }
                        catch
                        {
                            extras = new Dictionary<string, string>();
                        }
                    }

                    parsedRows.Add((row, extras));
                }

                using (var package = new ExcelPackage())
                {
                    var ws = package.Workbook.Worksheets.Add("Merge Report");

                    int col = 1;
                    foreach (var prop in baseProperties)
                    {
                        ws.Cells[1, col].Value = prop.Name;
                        ws.Cells[1, col].Style.Font.Bold = true;
                        col++;
                    }

                    var extraHeaderList = extraHeaders.OrderBy(x => x).ToList();

                    foreach (var key in extraHeaderList)
                    {
                        ws.Cells[1, col].Value = key;
                        ws.Cells[1, col].Style.Font.Bold = true;
                        col++;
                    }

                    int rowIdx = 2;

                    foreach (var (item, extras) in parsedRows)
                    {
                        col = 1;

                        foreach (var prop in baseProperties)
                            ws.Cells[rowIdx, col++].Value = prop.GetValue(item)?.ToString() ?? "";

                        foreach (var key in extraHeaderList)
                        {
                            extras.TryGetValue(key, out var val);
                            ws.Cells[rowIdx, col++].Value = val ?? "";
                        }

                        using (var range = ws.Cells[rowIdx, 1, rowIdx, col - 1])
                        {
                            range.Style.Fill.PatternType = ExcelFillStyle.Solid;

                            if (deletedRows.Any(x => x.Id == item.Id))
                                range.Style.Fill.BackgroundColor.SetColor(Color.LightCoral); // 🔴 old

                            else if (newRows.Any(x => x.Id == item.Id))
                                range.Style.Fill.BackgroundColor.SetColor(Color.LightGreen); // 🟢 new
                        }

                        rowIdx++;
                    }

                    ws.Cells[ws.Dimension.Address].AutoFitColumns();
                    ws.View.FreezePanes(2, 1);

                    // ✅ CLEAN SHEET (FINAL DATA)
                    var wsClean = package.Workbook.Worksheets.Add("Clean NRData");

                    col = 1;
                    foreach (var prop in baseProperties)
                    {
                        wsClean.Cells[1, col].Value = prop.Name;
                        wsClean.Cells[1, col].Style.Font.Bold = true;
                        col++;
                    }

                    var cleanRows = await _context.NRDatas
                        .Where(x => x.ProjectId == ProjectId && x.Status == true)
                        .ToListAsync();

                    int cleanRow = 2;

                    foreach (var item in cleanRows)
                    {
                        col = 1;
                        foreach (var prop in baseProperties)
                            wsClean.Cells[cleanRow, col++].Value = prop.GetValue(item)?.ToString() ?? "";

                        cleanRow++;
                    }

                    wsClean.Cells[wsClean.Dimension.Address].AutoFitColumns();
                    wsClean.View.FreezePanes(2, 1);

                    package.SaveAs(new FileInfo(filePath));
                }

                return Ok(new { MergedRows = mergedCount });
            }
            catch (Exception ex)
            {
                await _logger.LogErrorAsync("Error solving duplicates", ex.Message, nameof(DuplicateController));
                return StatusCode(500, "Internal server error");
            }
        }
        [HttpPost("Enhancement")]
        public async Task<IActionResult> ApplyEnhancement(int ProjectId)
        {
            try
            {
                // Normalize NULL numeric fields to 0 to avoid materialization errors
                await _context.Database.ExecuteSqlRawAsync(@"
UPDATE NRDatas
SET Quantity = IFNULL(Quantity, 0),
    NRQuantity = IFNULL(NRQuantity, 0),
    Pages = IFNULL(Pages, 0),
    RouteSort = IFNULL(RouteSort, 0),
    CenterSort = IFNULL(CenterSort, 0),
    NodalSort = IFNULL(NodalSort, 0),
    Steps = IFNULL(Steps, 0),
    LotNo = IFNULL(LotNo, 0)
WHERE ProjectId = {0};", ProjectId);

                var data = await _context.NRDatas
                    .Where(p => p.ProjectId == ProjectId && p.Status==true && p.Steps==1)
                    .ToListAsync();

                if (!data.Any())
                    return NotFound("Nr data not found for this project.");

                var projectconfig = await _context.ProjectConfigs
                    .Where(p => p.ProjectId == ProjectId).FirstOrDefaultAsync();

                if (projectconfig == null)
                {
                    return NotFound("Project config not exists for this project");
                }

                int smallestInner = 0;
                var innerEnv = projectconfig.Envelope;

                if (!string.IsNullOrEmpty(innerEnv))
                {
                    try
                    {
                        var envelopeDict = System.Text.Json.JsonSerializer.Deserialize<Dictionary<string, string>>(innerEnv);
                        if (envelopeDict != null && envelopeDict.TryGetValue("Inner", out var innerValue) &&
                            !string.IsNullOrWhiteSpace(innerValue))
                        {
                            var innerSizes = envelopeDict["Inner"]
                                .Split(',', StringSplitOptions.RemoveEmptyEntries)
                                .Select(e => e.Trim().ToUpper().Replace("E", ""))
                                .Where(x => int.TryParse(x, out _))
                                .Select(int.Parse)
                                .OrderBy(x => x)
                                .ToList();

                            if (innerSizes.Any())
                            {
                                smallestInner = innerSizes.First();
                            }
                        }
                        else if (envelopeDict != null && envelopeDict.TryGetValue("Outer", out var outerValue) &&
                                 !string.IsNullOrWhiteSpace(outerValue))
                        {
                            var outerSizes = envelopeDict["Outer"]
                                .Split(',', StringSplitOptions.RemoveEmptyEntries)
                                .Select(e => e.Trim().ToUpper().Replace("E", ""))
                                .Where(x => int.TryParse(x, out _))
                                .Select(int.Parse)
                                .OrderBy(x => x)
                                .ToList();

                            if (outerSizes.Any())
                            {
                                smallestInner = outerSizes.First();
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        await _logger.LogErrorAsync("Error in serializing Envelope", ex.Message, nameof(DuplicateController));
                        // Fallback: proceed without envelope-based rounding
                        smallestInner = 0;
                    }
                }

                // Consolidated calculation logic: Round Before (Optional) -> Enhance -> Round After (Mandatory if capacity exists)
                foreach (var d in data)
                {
                    if (d.NRQuantity > 0)
                    {
                        double initialQuantity = d.NRQuantity;

                        // Phase 1: Round before enhancement (if enabled)
                        if (projectconfig.RoundOffBeforeEnhancement && smallestInner > 0)
                        {
                            initialQuantity = Math.Ceiling(initialQuantity / (double)smallestInner) * smallestInner;
                        }

                        // Phase 2: Calculate enhancement value based on the initial quantity
                        double enhancementVal = 0;
                        if (projectconfig.Enhancement > 0)
                        {
                            enhancementVal = (projectconfig.Enhancement * initialQuantity) / 100.0;
                        }

                        // Phase 3: Add enhancement and apply final rounding
                        double totalTarget = initialQuantity + enhancementVal;

                        if (smallestInner > 0)
                        {
                            d.Quantity = (int)Math.Ceiling(totalTarget / (double)smallestInner) * smallestInner;
                        }
                        else
                        {
                            d.Quantity = (int)Math.Round(totalTarget);
                        }
                    }
                    d.Steps = Tools.Models.PipelineNavigator.GetNextStep(Tools.Models.PipelineNavigator.STEP_DUP_PARTIAL, projectconfig?.Modules);
                }

                await _context.SaveChangesAsync();
                var reportPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", ProjectId.ToString());
                if (!Directory.Exists(reportPath))
                    Directory.CreateDirectory(reportPath);

                var filePath = Path.Combine(reportPath, "EnhancementReport.xlsx");

                if (System.IO.File.Exists(filePath))
                    System.IO.File.Delete(filePath);

                var baseProperties = typeof(NRData).GetProperties(BindingFlags.Public | BindingFlags.Instance)
                    .Where(p => p.Name != "NRDatas" && p.Name != "ProjectId")
                    .ToList();

                using (var package = new ExcelPackage())
                {
                    var ws = package.Workbook.Worksheets.Add("Enhancement Report");

                    int col = 1;
                    foreach (var prop in baseProperties)
                    {
                        ws.Cells[1, col].Value = prop.Name;
                        ws.Cells[1, col].Style.Font.Bold = true;
                        col++;
                    }

                    int row = 2;

                    foreach (var item in data)
                    {
                        col = 1;

                        foreach (var prop in baseProperties)
                        {
                            ws.Cells[row, col++].Value = prop.GetValue(item)?.ToString();
                        }
                        row++;
                    }

                    ws.Cells[ws.Dimension.Address].AutoFitColumns();
                    ws.View.FreezePanes(2, 1);

                    package.SaveAs(new FileInfo(filePath));
                }

                // Logging
                var triggeredBy = LogHelper.GetTriggeredBy(User);

                await _logger.LogEventAsync(
                    "Enhancement applied with report",
                    "Enhancement",
                    triggeredBy,
                    ProjectId,
                    string.Empty,
                    LogHelper.ToJson(new { ProjectId })
                );

                return Ok(new
                {
                    EnhancementApplied = projectconfig.Enhancement,
                    ReportPath = filePath
                });
            }
            catch (Exception ex)
            {
                await _logger.LogErrorAsync("Error applying enhancement", ex.Message, nameof(DuplicateController));
                return StatusCode(500, "Internal server error");
            }
        }



        [HttpGet("MergeData")]
        public async Task<IActionResult> MergeEnvelope(int ProjectId)
        {
            // Helper function to parse JSON and sum the values (same as your existing one)
            int ParseJsonEnvelope(string value)
            {
                if (string.IsNullOrEmpty(value))
                {
                    return 0; // Return 0 if the value is null or empty
                }

                // Try parsing the string to an integer
                if (int.TryParse(value, out var result))
                {
                    return result;
                }
                else
                {
                    // Log invalid value if it cannot be parsed
                    Console.WriteLine($"Invalid value for envelope: {value}");
                    return 0;
                }
            }

            // Helper function to parse JSON and sum values for EnvelopeBreakages
            int ParseJsonEnvBreakage(string value)
            {
                if (string.IsNullOrEmpty(value))
                {
                    return 0; // Return 0 if the value is null or empty
                }

                // Sum values from JSON string like {\"E100\":\"1\",\"E50\":\"1\",\"E10\":\"3\"}
                try
                {
                    var envelopeData = Newtonsoft.Json.JsonConvert.DeserializeObject<Dictionary<string, int>>(value);
                    return envelopeData?.Values.Sum() ?? 0;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error parsing envelope data: {ex.Message}");
                    return 0;
                }
            }

            // Step 1: Group NRData by CatchNo and sum the quantity
            var NRdataGrouped = await _context.NRDatas
                .Where(p => p.ProjectId == ProjectId && p.Status==true)
                .GroupBy(p => p.CatchNo)
                .Select(g => new
                {
                    CatchNo = g.Key,
                    TotalQuantity = g.Sum(p => p.Quantity),
                    NRDataId = g.Select(p => p.Id).ToList()
                })
                .ToListAsync();

            // Step 2: Fetch ExtrasEnvelope data and then group/sum in-memory
            var extraEnvelopeData = await _context.ExtrasEnvelope
                .Where(p => p.ProjectId == ProjectId && p.Status == 1)
                .ToListAsync();

            var ExtraEnvGrouped = extraEnvelopeData
                .GroupBy(p => p.CatchNo)
                .Select(g =>
                {
                    // Sum the quantities and the envelope values
                    var totalQuantity = g.Sum(p => p.Quantity);
                    var innerEnvelopeSum = g.Sum(p => ParseJsonEnvelope(p.InnerEnvelope));
                    var outerEnvelopeSum = g.Sum(p => ParseJsonEnvelope(p.OuterEnvelope));

                    // Log the summed values
                    Console.WriteLine($"CatchNo: {g.Key}, TotalQuantity: {totalQuantity}, InnerEnvelopeSum: {innerEnvelopeSum}, OuterEnvelopeSum: {outerEnvelopeSum}");

                    return new
                    {
                        CatchNo = g.Key,
                        TotalQuantity = totalQuantity,
                        InnerEnvelopeSum = innerEnvelopeSum,
                        OuterEnvelopeSum = outerEnvelopeSum
                    };
                })
                .ToList();

            // Step 3: Fetch EnvelopeBreakages data and group by NRDataIds, then sum in-memory
            var innerEnvData = await _context.EnvelopeBreakages
                .Where(p => p.ProjectId == ProjectId)
                .ToListAsync();

            var processedInnerEnv = innerEnvData
                .GroupBy(p => p.NrDataId) // Grouping by NrDataId
                .Select(g => new
                {
                    NRDataId = g.Key,
                    TotalInnerEnvelopes = g.Sum(p => ParseJsonEnvBreakage(p.InnerEnvelope)),
                    TotalOuterEnvelopes = g.Sum(p => ParseJsonEnvBreakage(p.OuterEnvelope))
                })
                .ToList();

            // Step 4: Merge NRDataGrouped and processedInnerEnv in-memory
            var mergedResult = from nr in NRdataGrouped
                               join ee in ExtraEnvGrouped on nr.CatchNo equals ee.CatchNo into extraEnvGroup
                               from extraEnv in extraEnvGroup.DefaultIfEmpty()
                               select new
                               {
                                   CatchNo = nr.CatchNo,
                                   NRDataIds = nr.NRDataId,
                                   TotalNRDataQuantity = nr.TotalQuantity + (extraEnv?.TotalQuantity ?? 0),
                                   TotalInnerEnvEnvelopes = processedInnerEnv
                               .Where(env => nr.NRDataId.Contains(env.NRDataId))
                               .Sum(env => env.TotalInnerEnvelopes)
                               + (extraEnv?.InnerEnvelopeSum ?? 0),  // Add ExtraEnv inner envelope

                                   TotalOuterEnvEnvelopes = processedInnerEnv
                               .Where(env => nr.NRDataId.Contains(env.NRDataId))
                               .Sum(env => env.TotalOuterEnvelopes)
                               + (extraEnv?.OuterEnvelopeSum ?? 0)
                               };

            return Ok(mergedResult);
        }




    }
}
