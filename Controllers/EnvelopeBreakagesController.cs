using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using ERPToolsAPI.Data;
using Tools.Models;
using System.Text.Json;
using OfficeOpenXml.Style;
using OfficeOpenXml;
using System.Drawing;
using System.Reflection;
using Tools.Services;
using Microsoft.CodeAnalysis;
using static System.Runtime.InteropServices.JavaScript.JSType;
using System.Composition;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.Extensions.Options;

namespace Tools.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class EnvelopeBreakagesController : ControllerBase
    {
        private readonly ERPToolsDbContext _context;
        private readonly ILoggerService _loggerService;
        private readonly ApiSettings _apiSettings;

        public EnvelopeBreakagesController(ERPToolsDbContext context, ILoggerService loggerService, IOptions<ApiSettings> apiSettings)
        {
            _context = context;
            _loggerService = loggerService;
            _apiSettings = apiSettings.Value;
        }

        [HttpGet]
        public async Task<ActionResult> GetEnvelopeBreakages(int ProjectId, int? uploadId = null)
        {
            List<NRData> NRData;
            if (uploadId.HasValue)
            {
                var allData = await _context.NRDatas.Where(p => p.ProjectId == ProjectId).ToListAsync();
                NRData = allData.Where(p => p.UploadList != null && p.UploadList.Contains(uploadId.Value)).ToList();
            }
            else
            {
                NRData = await _context.NRDatas
                    .Where(p => p.ProjectId == ProjectId && p.Status == true)
                    .ToListAsync();
            }

            var Envelope = await _context.EnvelopeBreakages
                .Where(p => p.ProjectId == ProjectId)
                .ToListAsync();

            if (!NRData.Any() || !Envelope.Any())
                return NotFound("No data available for this project.");
            await _loggerService.LogEventAsync($"No data available for this project", "EnvelopeBreakage", LogHelper.GetTriggeredBy(User), ProjectId);


            var Consolidated = (from nr in NRData
                                join env in Envelope on nr.Id equals env.NrDataId
                                select new
                                {
                                    nr.Id,
                                    nr.ProjectId,
                                    nr.CourseName,
                                    nr.SubjectName,
                                    nr.NRDatas,
                                    nr.CatchNo,
                                    nr.CenterCode,
                                    nr.ExamTime,
                                    nr.ExamDate,
                                    nr.Quantity,
                                    nr.NodalCode,
                                    env.InnerEnvelope,
                                    env.OuterEnvelope
                                }).ToList();

            var reportPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", ProjectId.ToString());
            Directory.CreateDirectory(reportPath);

            var filename = uploadId.HasValue ? $"EnvelopeBreaking_v{uploadId}.xlsx" : $"EnvelopeBreaking.xlsx";
            var filePath = Path.Combine(reportPath, filename);

            // ?? Skip generation if file already exists
            if (System.IO.File.Exists(filePath))
            {
                return Ok(Consolidated); // Still return data for UI
            }

            // Collect unique keys from Inner/OuterEnvelope
            var innerKeys = new HashSet<string>();
            var outerKeys = new HashSet<string>();

            var parsedRows = new List<Dictionary<string, object>>();
            try
            {
                foreach (var row in Consolidated)
                {
                    var parsedRow = new Dictionary<string, object>
                    {
                        ["Id"] = row.Id,
                        ["ProjectId"] = row.ProjectId,
                        ["CourseName"] = row.CourseName,
                        ["SubjectName"] = row.SubjectName,
                        ["NRDatas"] = row.NRDatas,
                        ["CatchNo"] = row.CatchNo,
                        ["CenterCode"] = row.CenterCode,
                        ["ExamTime"] = row.ExamTime,
                        ["ExamDate"] = row.ExamDate,
                        ["Quantity"] = row.Quantity,
                        ["NodalCode"] = row.NodalCode,
                    };

                    // Parse NRDatas
                    if (!string.IsNullOrEmpty(row.NRDatas))
                    {
                        try
                        {
                            var nrDatasDict = JsonSerializer.Deserialize<Dictionary<string, string>>(row.NRDatas);
                            foreach (var kvp in nrDatasDict)
                                parsedRow[kvp.Key] = kvp.Value;
                        }
                        catch (Exception ex)
                        {
                            await _loggerService.LogErrorAsync("Error in NRDatas, not being able to serailize", ex.Message, nameof(EnvelopeBreakagesController));
                            return StatusCode(500, "Internal server error");
                        }
                    }

                    // Parse InnerEnvelope
                    if (!string.IsNullOrEmpty(row.InnerEnvelope))
                    {
                        try
                        {
                            var innerDict = JsonSerializer.Deserialize<Dictionary<string, string>>(row.InnerEnvelope);
                            foreach (var kvp in innerDict)
                            {
                                string key = $"Inner_{kvp.Key}";
                                parsedRow[key] = kvp.Value;
                                innerKeys.Add(key);
                            }
                        }
                        catch (Exception ex)
                        {
                            await _loggerService.LogErrorAsync("Error in InnerEnvelope, not being able to serailize", ex.Message, nameof(EnvelopeBreakagesController));
                            return StatusCode(500, "Internal server error");
                        }
                    }

                    // Parse OuterEnvelope
                    if (!string.IsNullOrEmpty(row.OuterEnvelope))
                    {
                        try
                        {
                            var outerDict = JsonSerializer.Deserialize<Dictionary<string, string>>(row.OuterEnvelope);
                            foreach (var kvp in outerDict)
                            {
                                string key = $"Outer_{kvp.Key}";
                                parsedRow[key] = kvp.Value;
                                outerKeys.Add(key);
                            }
                        }
                        catch (Exception ex)
                        {
                            await _loggerService.LogErrorAsync("Error in OuterEnvelope, not being able to serailize", ex.Message, nameof(EnvelopeBreakagesController));
                            return StatusCode(500, "Internal server error");
                        }
                    }

                    parsedRows.Add(parsedRow);
                    parsedRows.Add(parsedRow);

                }
            }
            catch (Exception ex)
            {
                await _loggerService.LogErrorAsync("Error in NRDatas, not being able to serailize", ex.Message, nameof(EnvelopeBreakagesController));
                return StatusCode(500, "Internal server error");
            }

            var allHeaders = parsedRows
                    .SelectMany(d => d.Keys)
                    .Union(innerKeys)
                    .Union(outerKeys)
                    .Distinct()
                    .OrderBy(k => k)
                    .ToList();

            var columnsToKeep = allHeaders.Where(header =>
            {
                return parsedRows.Any(rowDict =>
                    rowDict.ContainsKey(header) &&
                    !string.IsNullOrEmpty(rowDict[header]?.ToString()));
            }).ToList();

            // ?? Generate Excel report
            try
            {
                using (var package = new ExcelPackage())
                {
                    var ws = package.Workbook.Worksheets.Add("Envelope Report");

                    // Write headers
                    for (int i = 0; i < columnsToKeep.Count; i++)
                    {
                        ws.Cells[1, i + 1].Value = columnsToKeep[i];
                        ws.Cells[1, i + 1].Style.Font.Bold = true;
                    }

                    // Write data rows
                    int rowIdx = 2;
                    foreach (var rowDict in parsedRows)
                    {
                        int colIdx = 1;
                        foreach (var header in columnsToKeep)
                        {
                            rowDict.TryGetValue(header, out object value);
                            ws.Cells[rowIdx, colIdx++].Value = value?.ToString() ?? "";
                        }
                        rowIdx++;
                    }

                    ws.Cells[ws.Dimension.Address].AutoFitColumns();
                    ws.View.FreezePanes(2, 1);
                    package.SaveAs(new FileInfo(filePath));
                }
                await _loggerService.LogEventAsync($"EnvelopeBreakage report of ProjectId {ProjectId} has been created", "EnvelopeBreakage", LogHelper.GetTriggeredBy(User), ProjectId);
                return Ok(Consolidated); // Return original data for UI (optional)
            }
            catch (Exception ex)
            {

                await _loggerService.LogErrorAsync("Error in generating report", ex.Message, nameof(EnvelopeBreakagesController));
                return StatusCode(500, "Internal server error");
            }

        }


        // GET: api/EnvelopeBreakages/5
        [HttpGet("{id:int}")]
        public async Task<ActionResult<EnvelopeBreakage>> GetEnvelopeBreakage(int id)
        {
            var envelopeBreakage = await _context.EnvelopeBreakages.FindAsync(id);

            if (envelopeBreakage == null)
            {
                return NotFound();
            }

            return envelopeBreakage;
        }

        // PUT: api/EnvelopeBreakages/5
        // To protect from overposting attacks, see https://go.microsoft.com/fwlink/?linkid=2123754
        [HttpPut("{id}")]
        public async Task<IActionResult> PutEnvelopeBreakage(int id, EnvelopeBreakage envelopeBreakage)
        {
            if (id != envelopeBreakage.EnvelopeId)
            {
                return BadRequest();
            }

            _context.Entry(envelopeBreakage).State = EntityState.Modified;

            try
            {
                await _loggerService.LogEventAsync($"Updated EnvelopeBreakage with ID {id}", "EnvelopeBreakage", LogHelper.GetTriggeredBy(User), envelopeBreakage.ProjectId);
                await _context.SaveChangesAsync();
            }
            catch (Exception ex)
            {
                if (!EnvelopeBreakageExists(id))
                {
                    await _loggerService.LogEventAsync($"EnvelopeBreakage with ID {id} not found during updating", "EnvelopeBreakage", LogHelper.GetTriggeredBy(User), envelopeBreakage.ProjectId);
                    return NotFound();

                }
                else
                {
                    await _loggerService.LogErrorAsync("Error updating EnvelopeBreakage", ex.Message, nameof(EnvelopeBreakagesController));
                    return StatusCode(500, "Internal server error");
                }
            }

            return NoContent();
        }

        // POST: api/EnvelopeBreakages
        // To protect from overposting attacks, see https://go.microsoft.com/fwlink/?linkid=2123754
        [HttpPost("EnvelopeConfiguration")]
        public async Task<IActionResult> EnvelopeConfiguration(int ProjectId)
        {
            try
            {
                var envelopesJson = await _context.ProjectConfigs
                    .Where(s => s.ProjectId == ProjectId)
                    .Select(s => s.Envelope)
                    .FirstOrDefaultAsync();

                if (envelopesJson == null)
                    return NotFound("No envelope config found.");

                var nrDataList = await _context.NRDatas
                    .Where(s => s.ProjectId == ProjectId && s.Status == true)
                    .ToListAsync();

                if (!nrDataList.Any())
                    return NotFound("No NRData found.");
                var envelopeDict = JsonSerializer.Deserialize<Dictionary<string, string>>(envelopesJson);
                if (envelopeDict == null)
                    return BadRequest("Invalid envelope JSON format.");
                var innerSizes = envelopeDict["Inner"]
                    .Split(',', StringSplitOptions.RemoveEmptyEntries)
                    .Select(e => e.Trim().ToUpper().Replace("E", ""))
                    .Where(x => int.TryParse(x, out _))
                    .Select(int.Parse)
                    .OrderByDescending(x => x)
                    .ToList();

                var outerSizes = envelopeDict["Outer"]
                    .Split(',', StringSplitOptions.RemoveEmptyEntries)
                    .Select(e => e.Trim().ToUpper().Replace("E", ""))
                    .Where(x => int.TryParse(x, out _))
                    .Select(int.Parse)
                    .OrderByDescending(x => x)
                    .ToList();

                var env = await _context.EnvelopeBreakages.Where(p => p.ProjectId == ProjectId).ToListAsync();
                if (env.Any()) // Check if any records were found
                {
                    _context.EnvelopeBreakages.RemoveRange(env);
                    await _context.SaveChangesAsync();

                    await _loggerService.LogEventAsync($"Successfully deleted {env.Count} Envelope Breaking entries for ProjectID {ProjectId}", "EnvelopeBreakages",
                        LogHelper.GetTriggeredBy(User), ProjectId);
                }
                else
                {
                    // Handle the case where no records were found
                    Console.WriteLine("No breakages found for the specified project.");
                }
                var breakagesToAdd = new List<EnvelopeBreakage>();

                foreach (var row in nrDataList)
                {
                    int quantity = row.Quantity;

                    // ================= INNER =================
                    int remaining = quantity;
                    Dictionary<string, string> innerBreakdown = new();

                    foreach (var size in innerSizes)
                    {
                        int count = remaining / size;

                        if (count > 0)
                        {
                            innerBreakdown[$"E{size}"] = count.ToString();
                            remaining -= count * size;
                        }
                    }

                    // ================= OUTER =================
                    remaining = quantity;
                    Dictionary<string, string> outerBreakdown = new();
                    int totalOuterCount = 0;

                    // Step 1: Greedy distribution
                    foreach (var size in outerSizes)
                    {
                        int count = remaining / size;

                        if (count > 0)
                        {
                            outerBreakdown[$"E{size}"] = count.ToString();
                            totalOuterCount += count;
                            remaining -= count * size;
                        }
                    }

                    // Step 2: Handle leftover ONLY if inner != outer
                    bool sameEnvelopes = innerSizes.SequenceEqual(outerSizes);

                    if (remaining > 0 && !sameEnvelopes)
                    {
                        int smallestSize = outerSizes.Last();

                        int count = (int)Math.Ceiling((double)remaining / smallestSize);

                        if (outerBreakdown.ContainsKey($"E{smallestSize}"))
                        {
                            outerBreakdown[$"E{smallestSize}"] =
                                (int.Parse(outerBreakdown[$"E{smallestSize}"]) + count).ToString();
                        }
                        else
                        {
                            outerBreakdown[$"E{smallestSize}"] = count.ToString();
                        }

                        totalOuterCount += count;
                        remaining = 0;
                    }


                    var envelope = new EnvelopeBreakage
                    {
                        ProjectId = ProjectId,
                        NrDataId = row.Id,
                        InnerEnvelope = JsonSerializer.Serialize(innerBreakdown),
                        OuterEnvelope = JsonSerializer.Serialize(outerBreakdown),
                        TotalEnvelope = totalOuterCount
                    };

                    // ? Add to database

                    breakagesToAdd.Add(envelope);

                }

                if (breakagesToAdd.Any())
                {
                    _context.EnvelopeBreakages.AddRange(breakagesToAdd);
                    await _context.SaveChangesAsync();
                    await _loggerService.LogEventAsync($"Created Envelope Breaking of ProjectID {ProjectId}", "EnvelopeBreakages", LogHelper.GetTriggeredBy(User), ProjectId);
                }


                            // ✅ Call ProcessEnvelopeBreaking directly instead of via HTTP
                            try
                            {
                                // Create an Options wrapper for ApiSettings
                                var apiSettingsOptions = Options.Create(_apiSettings);
                                var envelopeBreakageProcessingController = new EnvelopeBreakageProcessingController(_context, _loggerService, apiSettingsOptions);
                                var processResult = await envelopeBreakageProcessingController.ProcessEnvelopeBreaking(ProjectId, LogHelper.GetTriggeredBy(User));
                                
                                if (processResult is OkObjectResult okResult)
                                {
                                    Console.WriteLine($"ProcessEnvelopeBreaking Success: {okResult.Value}");
                                }
                                else if (processResult is BadRequestObjectResult badResult)
                                {
                                    Console.WriteLine($"ProcessEnvelopeBreaking Failed: {badResult.Value}");
                                    return BadRequest(new { error = "Envelope breaking processing failed", details = badResult.Value });
                                }
                                else if (processResult is ObjectResult objResult && objResult.StatusCode >= 400)
                                {
                                    Console.WriteLine($"ProcessEnvelopeBreaking Error: {objResult.Value}");
                                    return StatusCode(objResult.StatusCode ?? 500, new { error = "Envelope breaking processing failed", details = objResult.Value });
                                }
                                else
                                {
                                    Console.WriteLine($"ProcessEnvelopeBreaking returned: {processResult.GetType().Name}");
                                }
                            }
                            catch (Exception ex)
                            {
                                var innerError = ex.InnerException?.Message ?? ex.Message;
                                Console.WriteLine($"ProcessEnvelopeBreaking Exception: {ex.Message} | Inner: {innerError}");
                                return StatusCode(500, new { 
                                    error = "Failed to process envelope breaking", 
                                    details = ex.Message,
                                    innerException = innerError,
                                    stackTrace = ex.StackTrace
                                });
                            }
                return Ok("Envelope breakdown report has been successfully created.");
            }
            catch (Exception ex)
            {
                await _loggerService.LogErrorAsync("Error creating EnvelopeBreakage", ex.Message, nameof(EnvelopeBreakagesController));
                return StatusCode(500, "Internal Server Error");
            }
        }

        [HttpGet("Reports/Exists")]
        public IActionResult CheckReportExists(int projectId, string fileName, int? uploadId = null)
        {
            var rootFolder = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", projectId.ToString());

            // If uploadId is provided and fileName doesn't already have a version, inject it
            string finalFileName = fileName;
            if (uploadId.HasValue)
            {
                string extension = Path.GetExtension(fileName);
                string nameWithoutExt = Path.GetFileNameWithoutExtension(fileName);
                if (!nameWithoutExt.Contains("_v"))
                {
                    finalFileName = $"{nameWithoutExt}_v{uploadId}{extension}";
                }
            }

            var filePath = Path.Combine(rootFolder, finalFileName);
            bool fileExists = System.IO.File.Exists(filePath);
            return Ok(new { exists = fileExists, fileName = finalFileName });
        }

        [HttpGet("EnvelopeSummaryReport")]
        public async Task<IActionResult> EnvelopeSummaryReport(int ProjectId, int? uploadId = null)
        {
            try
            {
                List<NRData> nrDataList;
                if (uploadId.HasValue)
                {
                    var allData = await _context.NRDatas.Where(p => p.ProjectId == ProjectId).ToListAsync();
                    nrDataList = allData.Where(p => p.UploadList != null && p.UploadList.Contains(uploadId.Value)).ToList();
                }
                else
                {
                    nrDataList = await _context.NRDatas
                        .Where(x => x.ProjectId == ProjectId && x.Status == true)
                        .ToListAsync();
                }

                if (!nrDataList.Any())
                    return NotFound("No NRData found.");

                if (!nrDataList.Any())
                    return NotFound("No NRData found.");

                var breakages = await _context.EnvelopeBreakages
                    .Where(x => x.ProjectId == ProjectId)
                    .ToListAsync();

                var extraEnvelopes = await _context.ExtrasEnvelope
                    .Where(x => x.ProjectId == ProjectId && x.Status == 1)
                    .ToListAsync();

                var extraConfigs = await _context.ExtraConfigurations
                    .Where(x => x.ProjectId == ProjectId)
                    .ToListAsync();

                if (!breakages.Any() && !extraEnvelopes.Any())
                    return NotFound("No EnvelopeBreakages or ExtraEnvelopes found.");

                // Collect all dynamic keys
                var allInnerKeys = new HashSet<string>();
                var allOuterKeys = new HashSet<string>();

                foreach (var b in breakages)
                {
                    if (!string.IsNullOrEmpty(b.InnerEnvelope))
                    {
                        try
                        {
                            var innerDict = JsonSerializer.Deserialize<Dictionary<string, string>>(b.InnerEnvelope);
                            if (innerDict != null)
                            {
                                foreach (var key in innerDict.Keys)
                                    allInnerKeys.Add($"Inner_{key}");
                            }
                        }
                        catch { }
                    }

                    if (!string.IsNullOrEmpty(b.OuterEnvelope))
                    {
                        try
                        {
                            var outerDict = JsonSerializer.Deserialize<Dictionary<string, string>>(b.OuterEnvelope);
                            if (outerDict != null)
                            {
                                foreach (var key in outerDict.Keys)
                                    allOuterKeys.Add($"Outer_{key}");
                            }
                        }
                        catch { }
                    }
                }

                // Collect keys from extra envelopes
                foreach (var extra in extraEnvelopes)
                {
                    var config = extraConfigs.FirstOrDefault(c => c.ExtraType == extra.ExtraId);
                    if (config != null)
                    {
                        try
                        {
                            var envelopeType = JsonSerializer.Deserialize<Dictionary<string, string>>(config.EnvelopeType);
                            if (envelopeType != null)
                            {
                                if (envelopeType.ContainsKey("Inner"))
                                    allInnerKeys.Add($"Inner_{envelopeType["Inner"]}");
                                if (envelopeType.ContainsKey("Outer"))
                                    allOuterKeys.Add($"Outer_{envelopeType["Outer"]}");
                            }
                        }
                        catch { }
                    }
                }

                using var package = new ExcelPackage();
                var worksheet = package.Workbook.Worksheets.Add("EnvelopeSummary");

                var nrProps = typeof(NRData).GetProperties().Select(p => p.Name).ToList();
                var headers = new List<string>();

                headers.AddRange(nrProps);
                headers.AddRange(allInnerKeys.OrderBy(x => x));
                headers.AddRange(allOuterKeys.OrderBy(x => x));
                headers.Add("TotalEnvelope");

                // Add headers
                for (int i = 0; i < headers.Count; i++)
                    worksheet.Cells[1, i + 1].Value = headers[i];

                using (var range = worksheet.Cells[1, 1, 1, headers.Count])
                {
                    range.Style.Font.Bold = true;
                    range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                }

                int rowIndex = 2;

                // Process NRData with EnvelopeBreakages
                foreach (var nr in nrDataList)
                {
                    var breakage = breakages.FirstOrDefault(x => x.NrDataId == nr.Id);

                    var rowDict = new Dictionary<string, object>();

                    foreach (var prop in typeof(NRData).GetProperties())
                        rowDict[prop.Name] = prop.GetValue(nr);

                    foreach (var key in allInnerKeys)
                        rowDict[key] = 0;

                    foreach (var key in allOuterKeys)
                        rowDict[key] = 0;

                    if (breakage != null)
                    {
                        try
                        {
                            var innerDict = JsonSerializer.Deserialize<Dictionary<string, string>>(breakage.InnerEnvelope ?? "{}");
                            if (innerDict != null)
                            {
                                foreach (var kvp in innerDict)
                                    rowDict[$"Inner_{kvp.Key}"] = kvp.Value;
                            }
                        }
                        catch { }

                        try
                        {
                            var outerDict = JsonSerializer.Deserialize<Dictionary<string, string>>(breakage.OuterEnvelope ?? "{}");
                            if (outerDict != null)
                            {
                                foreach (var kvp in outerDict)
                                    rowDict[$"Outer_{kvp.Key}"] = kvp.Value;
                            }
                        }
                        catch { }

                        rowDict["TotalEnvelope"] = breakage.TotalEnvelope;
                    }

                    for (int col = 0; col < headers.Count; col++)
                        worksheet.Cells[rowIndex, col + 1].Value =
                            rowDict.ContainsKey(headers[col]) ? rowDict[headers[col]] : 0;

                    rowIndex++;
                }

                // Process ExtraEnvelopes with different sort values
                foreach (var extra in extraEnvelopes)
                {
                    var baseRow = nrDataList.FirstOrDefault(x => x.CatchNo == extra.CatchNo);
                    if (baseRow == null)
                        continue;

                    var config = extraConfigs.FirstOrDefault(c => c.ExtraType == extra.ExtraId);
                    if (config == null)
                        continue;

                    var rowDict = new Dictionary<string, object>();

                    // Copy base NRData properties
                    foreach (var prop in typeof(NRData).GetProperties())
                    {
                        if (prop.Name == "Quantity")
                            rowDict[prop.Name] = extra.Quantity;
                        if (prop.Name == "NRQuantity")
                            rowDict[prop.Name] = "";
                        else
                            rowDict[prop.Name] = prop.GetValue(baseRow);
                    }
                    // Initialize all envelope keys
                    foreach (var key in allInnerKeys)
                        rowDict[key] = 0;

                    foreach (var key in allOuterKeys)
                        rowDict[key] = 0;

                    // Set CenterCode and sort values based on ExtraId
                    switch (extra.ExtraId)
                    {
                        case 1:
                            rowDict["CenterCode"] = "Nodal Extra";
                            rowDict["NodalSort"] = baseRow.NodalSort;
                            rowDict["CenterSort"] = 10000;
                            rowDict["RouteSort"] = baseRow.RouteSort;
                            break;
                        case 2:
                            rowDict["CenterCode"] = "University Extra";
                            rowDict["NodalSort"] = 10000;
                            rowDict["CenterSort"] = 100000;
                            rowDict["RouteSort"] = 10000;
                            break;
                        case 3:
                            rowDict["CenterCode"] = "Office Extra";
                            rowDict["NodalSort"] = 100000;
                            rowDict["CenterSort"] = 1000000;
                            rowDict["RouteSort"] = 100000;
                            break;
                        default:
                            rowDict["CenterCode"] = "Extra";
                            break;
                    }

                    // Add envelope data from config
                    try
                    {
                        var envelopeType = JsonSerializer.Deserialize<Dictionary<string, string>>(config.EnvelopeType);
                        if (envelopeType != null)
                        {
                            if (envelopeType.ContainsKey("Inner") && !string.IsNullOrEmpty(extra.InnerEnvelope))
                            {
                                string innerKey = $"Inner_{envelopeType["Inner"]}";
                                rowDict[innerKey] = extra.InnerEnvelope;
                            }

                            if (envelopeType.ContainsKey("Outer") && !string.IsNullOrEmpty(extra.OuterEnvelope))
                            {
                                string outerKey = $"Outer_{envelopeType["Outer"]}";
                                rowDict[outerKey] = extra.OuterEnvelope;
                            }
                        }
                    }
                    catch { }

                    // Calculate TotalEnvelope as sum of outer counts
                    int totalOuter = 0;
                    foreach (var key in allOuterKeys)
                    {
                        if (rowDict.ContainsKey(key) && int.TryParse(rowDict[key].ToString(), out var val))
                            totalOuter += val;
                    }
                    rowDict["TotalEnvelope"] = totalOuter;

                    for (int col = 0; col < headers.Count; col++)
                        worksheet.Cells[rowIndex, col + 1].Value =
                            rowDict.ContainsKey(headers[col]) ? rowDict[headers[col]] : 0;

                    rowIndex++;
                }

                if (worksheet.Dimension != null)
                {
                    worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
                    worksheet.View.FreezePanes(2, 1);
                }
                // Save in application root folder
                var reportPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", ProjectId.ToString());
                Directory.CreateDirectory(reportPath);
                var fileName = uploadId.HasValue ? $"EnvelopeSummary_v{uploadId}.xlsx" : "EnvelopeSummary.xlsx";
                var filePath = Path.Combine(reportPath, fileName);
                if (System.IO.File.Exists(filePath))
                {
                    System.IO.File.Delete(filePath);
                }

                package.SaveAs(new FileInfo(filePath));

                return Ok(new { message = $"Envelope summary report saved at root folder: {filePath}", fileName });

            }
            catch (Exception ex)
            {
                await _loggerService.LogErrorAsync("Error generating EnvelopeSummaryReport", ex.ToString(), nameof(EnvelopeBreakagesController));
                return StatusCode(500, $"Internal Server Error: {ex.Message}");
            }
        }


        [HttpGet("CatchEnvelopeSummaryWithExtras")]
        public async Task<IActionResult> CatchEnvelopeSummaryWithExtras(int ProjectId, int? uploadId = null)
        {
            try
            {
                // ==============================
                // 1️⃣ Get NRData
                // ==============================
                List<NRData> nrDataList;
                if (uploadId.HasValue)
                {
                    var allData = await _context.NRDatas.Where(p => p.ProjectId == ProjectId).ToListAsync();
                    nrDataList = allData.Where(p => p.UploadList != null && p.UploadList.Contains(uploadId.Value)).ToList();
                }
                else
                {
                    nrDataList = await _context.NRDatas
                        .Where(x => x.ProjectId == ProjectId && x.Status == true)
                        .ToListAsync();
                }

                if (!nrDataList.Any())
                    return NotFound("No NRData found.");

                var nrDataMap = nrDataList
                    .GroupBy(x => x.CatchNo)
                    .ToDictionary(g => g.Key, g => g.First());

                var uniqueFields = await _context.Fields
                    .Where(f => f.IsUnique)
                    .Select(f => f.Name)
                    .ToListAsync();

                // ✅ Distinct Nodal count per CatchNo
                var nodalCountMap = nrDataList
                    .Where(x => !string.IsNullOrEmpty(x.CatchNo))
                    .GroupBy(x => x.CatchNo)
                    .ToDictionary(
                        g => g.Key,
                        g => g.Select(x => x.NodalCode).Distinct().Count()
                    );

                // ==============================
                // 2️⃣ EnvelopeBreakages
                // ==============================
                var breakages = await _context.EnvelopeBreakages
                    .Where(x => x.ProjectId == ProjectId)
                    .ToListAsync();

                // ==============================
                // 3️⃣ Extra Envelopes + Config
                // ==============================
                var extraEnvelopes = await _context.ExtrasEnvelope
                    .Where(x => x.ProjectId == ProjectId && x.Status == 1)
                    .ToListAsync();

                var extraConfigs = await _context.ExtraConfigurations
                    .Where(x => x.ProjectId == ProjectId)
                    .ToListAsync();

                // 🔥 Fast lookup
                var configMap = extraConfigs.ToDictionary(c => c.ExtraType, c => c);

                var summaryData = new Dictionary<string, Dictionary<string, int>>();
                var qtyData = new Dictionary<string, int>();

                var allInnerKeys = new HashSet<string>();
                var allOuterKeys = new HashSet<string>();

                // ==============================
                // 4️⃣ Process Breakages
                // ==============================
                foreach (var nr in nrDataList)
                {
                    string catchNo = nr.CatchNo;
                    if (string.IsNullOrEmpty(catchNo)) continue;

                    if (!summaryData.ContainsKey(catchNo))
                        summaryData[catchNo] = new Dictionary<string, int>();

                    if (!qtyData.ContainsKey(catchNo))
                        qtyData[catchNo] = 0;

                    qtyData[catchNo] += nr.Quantity;

                    var breakage = breakages.FirstOrDefault(x => x.NrDataId == nr.Id);

                    if (breakage == null) continue;

                    // INNER
                    if (!string.IsNullOrEmpty(breakage.InnerEnvelope))
                    {
                        try
                        {
                            var dict = JsonSerializer.Deserialize<Dictionary<string, string>>(breakage.InnerEnvelope);
                            if (dict != null)
                            {
                                foreach (var kvp in dict)
                                {
                                    string key = $"Inner_{kvp.Key}";
                                    allInnerKeys.Add(key);

                                    int val = int.TryParse(kvp.Value, out var v) ? v : 0;

                                    if (!summaryData[catchNo].ContainsKey(key))
                                        summaryData[catchNo][key] = 0;

                                    summaryData[catchNo][key] += val;
                                }
                            }
                        }
                        catch { }
                    }

                    // OUTER
                    if (!string.IsNullOrEmpty(breakage.OuterEnvelope))
                    {
                        try
                        {
                            var dict = JsonSerializer.Deserialize<Dictionary<string, string>>(breakage.OuterEnvelope);
                            if (dict != null)
                            {
                                foreach (var kvp in dict)
                                {
                                    string key = $"Outer_{kvp.Key}";
                                    allOuterKeys.Add(key);

                                    int val = int.TryParse(kvp.Value, out var v) ? v : 0;

                                    if (!summaryData[catchNo].ContainsKey(key))
                                        summaryData[catchNo][key] = 0;

                                    summaryData[catchNo][key] += val;
                                }
                            }
                        }
                        catch { }
                    }
                }

                // ==============================
                // 5️⃣ Process Extra Envelopes
                // ==============================
                foreach (var extra in extraEnvelopes)
                {
                    string catchNo = extra.CatchNo;
                    if (string.IsNullOrEmpty(catchNo)) continue;

                    if (!summaryData.ContainsKey(catchNo))
                        summaryData[catchNo] = new Dictionary<string, int>();

                    if (!qtyData.ContainsKey(catchNo))
                        qtyData[catchNo] = 0;

                    // ✅ Multiplier logic
                    int multiplier = 1;
                    if (extra.ExtraId == 1 && nodalCountMap.ContainsKey(catchNo))
                    {
                        multiplier = nodalCountMap[catchNo];
                    }

                    qtyData[catchNo] += extra.Quantity * multiplier;

                    if (!configMap.TryGetValue(extra.ExtraId, out var config))
                        continue;

                    try
                    {
                        var envelopeType = JsonSerializer.Deserialize<Dictionary<string, string>>(config.EnvelopeType);
                        if (envelopeType == null) continue;

                        // INNER
                        if (envelopeType.ContainsKey("Inner") && !string.IsNullOrEmpty(extra.InnerEnvelope))
                        {
                            string key = $"Inner_{envelopeType["Inner"]}";
                            allInnerKeys.Add(key);

                            int val = int.TryParse(extra.InnerEnvelope, out var v) ? v : 0;
                            val *= multiplier;

                            if (!summaryData[catchNo].ContainsKey(key))
                                summaryData[catchNo][key] = 0;

                            summaryData[catchNo][key] += val;
                        }

                        // OUTER
                        if (envelopeType.ContainsKey("Outer") && !string.IsNullOrEmpty(extra.OuterEnvelope))
                        {
                            string key = $"Outer_{envelopeType["Outer"]}";
                            allOuterKeys.Add(key);

                            int val = int.TryParse(extra.OuterEnvelope, out var v) ? v : 0;
                            val *= multiplier;

                            if (!summaryData[catchNo].ContainsKey(key))
                                summaryData[catchNo][key] = 0;

                            summaryData[catchNo][key] += val;
                        }
                    }
                    catch { }
                }

                // ==============================
                // 6️⃣ Headers
                // ==============================
                var allHeaders = summaryData
                    .SelectMany(x => x.Value.Keys)
                    .Distinct()
                    .ToList();

                var orderedHeaders = new List<string>();

                orderedHeaders.AddRange(uniqueFields);
                orderedHeaders.Add("CatchNo");
                orderedHeaders.Add("Qty");

                orderedHeaders.AddRange(allHeaders.Where(x => x.StartsWith("Inner_")).OrderBy(x => x));
                orderedHeaders.AddRange(allHeaders.Where(x => x.StartsWith("Outer_")).OrderBy(x => x));

                orderedHeaders.Add("TotalEnvelope");

                // ==============================
                // 7️⃣ Excel
                // ==============================
                using var package = new ExcelPackage();
                var ws = package.Workbook.Worksheets.Add("CatchSummary");

                for (int i = 0; i < orderedHeaders.Count; i++)
                    ws.Cells[1, i + 1].Value = orderedHeaders[i];

                ws.Cells[1, 1, 1, orderedHeaders.Count].Style.Font.Bold = true;

                int row = 2;

                foreach (var catchEntry in summaryData.OrderBy(x => x.Key))
                {
                    var catchNo = catchEntry.Key;
                    var nrRow = nrDataMap.ContainsKey(catchNo) ? nrDataMap[catchNo] : null;

                    int col = 1;

                    foreach (var field in uniqueFields)
                    {
                        var prop = typeof(NRData).GetProperty(field);
                        ws.Cells[row, col++].Value = (nrRow != null && prop != null) ? prop.GetValue(nrRow) : null;
                    }

                    ws.Cells[row, col++].Value = catchNo;
                    ws.Cells[row, col++].Value = qtyData.ContainsKey(catchNo) ? qtyData[catchNo] : 0;

                    int totalOuter = 0;

                    foreach (var key in orderedHeaders.Skip(uniqueFields.Count + 2).Take(orderedHeaders.Count - (uniqueFields.Count + 3)))
                    {
                        int val = catchEntry.Value.ContainsKey(key) ? catchEntry.Value[key] : 0;
                        ws.Cells[row, col++].Value = val;

                        if (key.StartsWith("Outer_"))
                            totalOuter += val;
                    }

                    ws.Cells[row, col].Value = totalOuter;

                    row++;
                }

                ws.Cells[ws.Dimension.Address].AutoFitColumns();
                ws.View.FreezePanes(2, 1);

                // ==============================
                // 8️⃣ Save
                // ==============================
                var folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", ProjectId.ToString());
                Directory.CreateDirectory(folderPath);

                var fileName = uploadId.HasValue ? $"CatchSummary_v{uploadId}.xlsx" : "CatchSummary.xlsx";
                var filePath = Path.Combine(folderPath, fileName);

                if (System.IO.File.Exists(filePath))
                    System.IO.File.Delete(filePath);

                package.SaveAs(new FileInfo(filePath));

                await _loggerService.LogEventAsync("CatchSummary report created", "CatchSummary", LogHelper.GetTriggeredBy(User), ProjectId);

                return Ok(new { message = $"CatchSummary.xlsx generated at: {filePath}", fileName });
            }
            catch (Exception ex)
            {
                await _loggerService.LogErrorAsync("Error generating CatchEnvelopeSummaryWithExtras", ex.ToString(), nameof(EnvelopeBreakagesController));
                return StatusCode(500, $"Internal Server Error: {ex.Message}");
            }
        }

        // DELETE: api/EnvelopeBreakages/5
        [HttpDelete("{id}")]
        public async Task<IActionResult> DeleteEnvelopeBreakage(int id)
        {
            try
            {
                var envelopeBreakage = await _context.EnvelopeBreakages.FindAsync(id);
                if (envelopeBreakage == null)
                {
                    return NotFound();
                }

                _context.EnvelopeBreakages.Remove(envelopeBreakage);
                await _loggerService.LogEventAsync($"Deleted Envelope Breaking of Id {id}", "EnvelopeBreakages", LogHelper.GetTriggeredBy(User), envelopeBreakage.ProjectId);
                await _context.SaveChangesAsync();

                return NoContent();
            }
            catch (Exception ex)
            {
                await _loggerService.LogErrorAsync("Error deleting EnvelopeBreakage", ex.Message, nameof(EnvelopeBreakagesController));
                return StatusCode(500, "Internal Server Error");
            }

        }

        private bool EnvelopeBreakageExists(int id)
        {
            return _context.EnvelopeBreakages.Any(e => e.EnvelopeId == id);
        }
    }
}

