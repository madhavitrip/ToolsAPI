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

        // GET: api/EnvelopeBreakages
        [HttpGet]
        public async Task<ActionResult> GetEnvelopeBreakages(int ProjectId)
        {
            var NRData = await _context.NRDatas
                .Where(p => p.ProjectId == ProjectId && p.Status == true)
                .ToListAsync();

            var Envelope = await _context.EnvelopeBreakages
                .Where(p => p.ProjectId == ProjectId)
                .ToListAsync();

            if (!NRData.Any() || !Envelope.Any())
                return NotFound("No data available for this project.");
            _loggerService.LogEvent($"No data available for this project", "EnvelopeBreakage", LogHelper.GetTriggeredBy(User), ProjectId);


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

            var filename = $"EnvelopeBreaking.xlsx";
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
                            _loggerService.LogError("Error in NRDatas, not being able to serailize", ex.Message, nameof(EnvelopeBreakagesController));
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
                            _loggerService.LogError("Error in InnerEnvelope, not being able to serailize", ex.Message, nameof(EnvelopeBreakagesController));
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
                            _loggerService.LogError("Error in OuterEnvelope, not being able to serailize", ex.Message, nameof(EnvelopeBreakagesController));
                            return StatusCode(500, "Internal server error");
                        }
                    }

                    parsedRows.Add(parsedRow);
                    parsedRows.Add(parsedRow);

                }
            }
            catch (Exception ex)
            {
                _loggerService.LogError("Error in NRDatas, not being able to serailize", ex.Message, nameof(EnvelopeBreakagesController));
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
                _loggerService.LogEvent($"EnvelopeBreakage report of ProjectId {ProjectId} has been created", "EnvelopeBreakage", LogHelper.GetTriggeredBy(User), ProjectId);
                return Ok(Consolidated); // Return original data for UI (optional)
            }
            catch (Exception ex)
            {

                _loggerService.LogError("Error in generating report", ex.Message, nameof(EnvelopeBreakagesController));
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
                _loggerService.LogEvent($"Updated EnvelopeBreakage with ID {id}", "EnvelopeBreakage", LogHelper.GetTriggeredBy(User), envelopeBreakage.ProjectId);
                await _context.SaveChangesAsync();
            }
            catch (Exception ex)
            {
                if (!EnvelopeBreakageExists(id))
                {
                    _loggerService.LogEvent($"EnvelopeBreakage with ID {id} not found during updating", "EnvelopeBreakage", LogHelper.GetTriggeredBy(User), envelopeBreakage.ProjectId);
                    return NotFound();

                }
                else
                {
                    _loggerService.LogError("Error updating EnvelopeBreakage", ex.Message, nameof(EnvelopeBreakagesController));
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

                // 🔥 Delete old records
                var env = await _context.EnvelopeBreakages
                    .Where(p => p.ProjectId == ProjectId)
                    .ToListAsync();

                if (env.Any())
                {
                    _context.EnvelopeBreakages.RemoveRange(env);
                    await _context.SaveChangesAsync();

                    _loggerService.LogEvent(
                        $"Deleted all Envelope Breaking entries for ProjectID {ProjectId}",
                        "EnvelopeBreakages",
                        LogHelper.GetTriggeredBy(User),
                        ProjectId);
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

                    breakagesToAdd.Add(envelope);
                }

                if (breakagesToAdd.Any())
                {
                    _context.EnvelopeBreakages.AddRange(breakagesToAdd);
                    await _context.SaveChangesAsync();

                    _loggerService.LogEvent(
                        $"Created Envelope Breaking of ProjectID {ProjectId}",
                        "EnvelopeBreakages",
                        LogHelper.GetTriggeredBy(User),
                        ProjectId);
                }

                // 🔗 API Call
                using var client = new HttpClient();
                var response = await client.PostAsync(
                    $"{_apiSettings.EnvelopeBreakageUrl}?ProjectId={ProjectId}&triggeredBy={LogHelper.GetTriggeredBy(User)}",
                    new StringContent("")
                );

                if (!response.IsSuccessStatusCode)
                {
                    var error = await response.Content.ReadAsStringAsync();
                    Console.WriteLine($"API Failed: {response.StatusCode}, {error}");
                }
                else
                {
                    var data = await response.Content.ReadAsStringAsync();
                    Console.WriteLine($"API Success: {data}");
                }

                return Ok("Envelope breakdown report has been successfully created.");
            }
            catch (Exception ex)
            {
                _loggerService.LogError("Error creating EnvelopeBreakage", ex.Message, nameof(EnvelopeBreakagesController));
                return StatusCode(500, "Internal Server Error");
            }
        }

        [HttpGet("Reports/Exists")]
        public IActionResult CheckReportExists(int projectId, string fileName)
        {
            var rootFolder = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", projectId.ToString());
            Console.WriteLine(rootFolder);
            var filePath = Path.Combine(rootFolder, fileName);  // Add the file name to the path
            Console.WriteLine(filePath);
            // Check if the file exists and return the result
            bool fileExists = System.IO.File.Exists(filePath);
            return Ok(new { exists = fileExists });
        }

        [HttpGet("EnvelopeSummaryReport")]
        public async Task<IActionResult> EnvelopeSummaryReport(int ProjectId)
        {
            try
            {
                var nrDataList = await _context.NRDatas
                    .Where(x => x.ProjectId == ProjectId && x.Status == true)
                    .ToListAsync();

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
                var filePath = Path.Combine(reportPath, "EnvelopeSummary.xlsx");
                if (System.IO.File.Exists(filePath))
                {
                    System.IO.File.Delete(filePath);
                }

                package.SaveAs(new FileInfo(filePath));

                return Ok($"Envelope summary report saved at root folder: {filePath}");

            }
            catch (Exception ex)
            {
                _loggerService.LogError("Error generating EnvelopeSummaryReport", ex.ToString(), nameof(EnvelopeBreakagesController));
                return StatusCode(500, $"Internal Server Error: {ex.Message}");
            }
        }


        [HttpGet("CatchEnvelopeSummaryWithExtras")]
        public async Task<IActionResult> CatchEnvelopeSummaryWithExtras(int ProjectId)
        {
            try
            {
                // ==============================
                // 1️⃣ Get NRData
                // ==============================
                var nrDataList = await _context.NRDatas
                    .Where(x => x.ProjectId == ProjectId && x.Status == true)
                    .ToListAsync();

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

                var filePath = Path.Combine(folderPath, "CatchSummary.xlsx");

                if (System.IO.File.Exists(filePath))
                    System.IO.File.Delete(filePath);

                package.SaveAs(new FileInfo(filePath));

                _loggerService.LogEvent("CatchSummary report created", "CatchSummary", LogHelper.GetTriggeredBy(User), ProjectId);

                return Ok($"CatchSummary.xlsx generated at: {filePath}");
            }
            catch (Exception ex)
            {
                _loggerService.LogError("Error generating CatchEnvelopeSummaryWithExtras", ex.ToString(), nameof(EnvelopeBreakagesController));
                return StatusCode(500, $"Internal Server Error: {ex.Message}");
            }
        }

        /* [HttpGet("Replication")]
         public async Task<IActionResult> ReplicationConfiguration(int ProjectId)
         {
             // Envelope capacities map (can be pulled from DB if needed)
             var envCaps = await _context.EnvelopesTypes
           .Select(e => new { e.EnvelopeName, e.Capacity })
           .ToListAsync();
             // Convert to dictionary for fast lookup
             Dictionary<string, int> envelopeCapacities = envCaps
                 .ToDictionary(x => x.EnvelopeName, x => x.Capacity);

             var nrData = await _context.NRDatas
                 .Where(p => p.ProjectId == ProjectId && p.Status == true)
                 .ToListAsync();
             var missingPages = nrData
          .Where(x => x.Pages == null || x.Pages <= 0)
          .Select(x => new { x.CatchNo, x.CenterCode })
          .ToList();

             if (missingPages.Any())
             {
                 var sample = missingPages
                     .Take(10)
                     .Select(x => $"{x.CatchNo}-{x.CenterCode}");

                 return BadRequest(new
                 {
                     message = "Pages are missing for some NRData. Please upload pages for all NRData before running replication.",
                     affectedRecords = sample
                 });
             }
             var ProjectConfig = await _context.ProjectConfigs
                 .Where(p => p.ProjectId == ProjectId).FirstOrDefaultAsync();

             var Boxcapacity = ProjectConfig.BoxCapacity;
             var OuterEnv = ProjectConfig.Envelope;
             var envelopeObj = JsonSerializer.Deserialize<Dictionary<string, string>>(ProjectConfig.Envelope);
             if (envelopeObj == null || !envelopeObj.ContainsKey("Outer"))
                 throw new Exception("Outer envelope configuration is missing in ProjectConfig.Envelope.");

             string outerEnvValue = envelopeObj["Outer"];
             if (string.IsNullOrWhiteSpace(outerEnvValue))
                 throw new Exception("Outer envelope value is empty in ProjectConfig.Envelope.");
             int envelopeSize = 0;
             // ?? Extract numeric part strictly
             var outerParts = outerEnvValue.Split(',', StringSplitOptions.RemoveEmptyEntries);

             // ?? Only parse if there�s a single non-empty value
             if (outerParts.Length == 1)
             {
                 string singleValue = outerParts[0].Trim();

                 // Extract numeric part (e.g., from "E10" -> "10")
                 string digits = new string(singleValue.Where(char.IsDigit).ToArray());

                 if (int.TryParse(digits, out int parsedValue) && parsedValue > 0)
                 {
                     envelopeSize = parsedValue;
                 }
             }
             else
             {
                 // Multiple envelopes, skip parsing (maybe log or handle separately)
                 envelopeSize = 0; // or keep default
             }

             var capacity = await _context.BoxCapacity
            .Where(c => c.BoxCapacityId == Boxcapacity)
              .Select(c => c.Capacity) // assuming the column is named 'Value'
           .FirstOrDefaultAsync();

             var boxIds = ProjectConfig.BoxBreakingCriteria;
             var sortingId = ProjectConfig.SortingBoxReport;
             var duplicatesFields = ProjectConfig.DuplicateRemoveFields;
             bool InnerBundling = ProjectConfig.IsInnerBundlingDone;
             var innerBundlingFieldNames = new List<string>();
             if (InnerBundling)
             {
                 var InnerBCriteria = ProjectConfig.InnerBundlingCriteria;
                 innerBundlingFieldNames = await _context.Fields
                   .Where(f => InnerBCriteria.Contains(f.FieldId))
                 .Select(f => f.Name)
                 .ToListAsync();
             }
             var fields = await _context.Fields
                 .Where(f => boxIds.Contains(f.FieldId))
                 .ToListAsync();
             var fieldsFromDb = await _context.Fields
            .Where(f => sortingId.Contains(f.FieldId))
            .ToListAsync();
             // Step 2: Fetch the corresponding field names from the Fields table
             var fieldNames = fieldsFromDb
                  .OrderBy(f => sortingId.IndexOf(f.FieldId))
                 .Select(f => f.Name)  // Get the field names
             .ToList();

             _loggerService.LogEvent($"Fieldnames  {string.Join(", ", fieldNames)}", "EnvelopeBreakages", LogHelper.GetTriggeredBy(User), ProjectId);
             var dupNames = await _context.Fields
                 .Where(f => duplicatesFields.Contains(f.FieldId))
                 .Select(f => f.Name)
                 .ToListAsync();

             var startBox = ProjectConfig.BoxNumber;
             bool resetOnSymbolChange = ProjectConfig.ResetOnSymbolChange;
             var extrasconfig = await _context.ExtraConfigurations
                 .Where(p => p.ProjectId == ProjectId)
                 .ToListAsync();

             var reportPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", ProjectId.ToString());
             if (!Directory.Exists(reportPath))
             {
                 Directory.CreateDirectory(reportPath);
             }

             var filename = "BoxBreaking.xlsx";
             var filePath = Path.Combine(reportPath, filename);

             // ?? Skip generation if file already exists
             if (System.IO.File.Exists(filePath))
             {
                 System.IO.File.Delete(filePath);
             }
             // Define path to breakingreport.xlsx
             var breakingReportPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", ProjectId.ToString(), "EnvelopeBreaking.xlsx");

             // Check if file exists
             if (!System.IO.File.Exists(breakingReportPath))
             {
                 return NotFound(new { message = "envelopebreakingreport.xlsx not found" });
             }

             var breakingReportData = new List<ExcelInputRow>();
             try
             {
                 using (var package = new ExcelPackage(new FileInfo(breakingReportPath)))
                 {
                     var worksheet = package.Workbook.Worksheets.FirstOrDefault();
                     if (worksheet == null)
                     {
                         return BadRequest(new { message = "No worksheet found in enevlopebreakingreport.xlsx" });
                     }

                     int rowCount = worksheet.Dimension.Rows;
                     for (int row = 2; row <= rowCount; row++)
                     {
                         try
                         {
                             var inputRow = new ExcelInputRow
                             {
                                 CatchNo = worksheet.Cells[row, 2].Text.Trim(),
                                 CenterCode = worksheet.Cells[row, 3].Text.Trim(),
                                CenterSort= Convert.ToInt32(worksheet.Cells[row, 4].Text.Trim()),
                                 ExamTime = worksheet.Cells[row, 15].Text.Trim(),
                                 ExamDate = worksheet.Cells[row, 16].Text.Trim(),
                                 Quantity = int.Parse(worksheet.Cells[row, 5].Text),
                                 TotalEnv = int.Parse(worksheet.Cells[row, 8].Text),
                                 NRQuantity = int.Parse(worksheet.Cells[row, 10].Text),
                                 NodalCode = worksheet.Cells[row, 11].Text.Trim(), 
                                 NodalSort = double.TryParse(worksheet.Cells[row, 12].Text.Trim(), out double sortVal)
                                 ? sortVal
                                 : 0.0,
                                 Route = worksheet.Cells[row, 13].Text.Trim(),
                                 RouteSort = Convert.ToInt32(worksheet.Cells[row, 14].Text.Trim()),
                                 OmrSerial = worksheet.Cells[row, 17].Text.Trim(), 
                                 CourseName = worksheet.Cells[row,18].Text.Trim(),
                             };

                             breakingReportData.Add(inputRow);
                         }
                         catch (Exception ex)
                         {
                             _loggerService.LogError($"Error parsing row {row}", ex.Message, nameof(EnvelopeBreakagesController));
                         }
                     }
                 }
             }
             catch (Exception ex)
             {
                 _loggerService.LogError($"Error in Breaking Report", ex.Message, nameof(EnvelopeBreakagesController));
             }


             // Step 1: Remove duplicates (CatchNo + CenterCode), preserving first occurrence
             var uniqueRows = breakingReportData
            .GroupBy(x =>
            {
             // Build a composite key using the fields listed in dupNames
             var keyParts = dupNames.Select(fieldName =>
             {
              var prop = x.GetType().GetProperty(fieldName, BindingFlags.IgnoreCase | BindingFlags.Public | BindingFlags.Instance);
              return prop?.GetValue(x)?.ToString()?.Trim() ?? string.Empty;
              });

             // Join them with an underscore (or any separator)
              return string.Join("_", keyParts);
               })
            .Select(g => g.First())
            .ToList();

             // Step 2: Calculate Start, End, Serial (before sorting)
             var enrichedList = new List<dynamic>();
             string previousCatchNo = null;
             int previousEnd = 0;
             try
             {
                 foreach (var row in uniqueRows)
                 {
                     int start = row.CatchNo != previousCatchNo ? 1 : previousEnd + 1;
                     int end = start + row.TotalEnv - 1;
                     string serial = $"{start} to {end}";

                     enrichedList.Add(new
                     {
                         row.CatchNo,
                         row.CenterCode,
                         row.CenterSort,
                         row.ExamTime,
                         row.ExamDate,
                         row.Quantity,
                         row.TotalEnv,
                         row.NRQuantity,
                         row.NodalCode,
                         row.NodalSort,
                         row.Route,
                         row.RouteSort,
                         Start = start,
                         End = end,
                         Serial = serial,
                         row.OmrSerial,
                         row.CourseName,
                     });

                     previousCatchNo = row.CatchNo;
                     previousEnd = end;
                 }
             }
             catch (Exception ex)
             {
                 _loggerService.LogError($"Error in UniqueRows", ex.Message, nameof(EnvelopeBreakagesController));
             }

             if (!enrichedList.Any())
                 return Ok(new { message = "No data to process" });
             // Step 1: Cache property info once
             var properties = fieldNames
                 .Select(name => new
                 {
                     Name = name,
                     Property = enrichedList.First().GetType().GetProperty(name,
     BindingFlags.Public | BindingFlags.Instance | BindingFlags.IgnoreCase)
                 })
                 .Where(x => x.Property != null)
                 .ToList();

             // Step 2: Apply ordering using cached properties
             IOrderedEnumerable<dynamic> ordered = null;

             for (int i = 0; i < properties.Count; i++)
             {
                 var prop = properties[i].Property;

                 Func<dynamic, object> keySelector = x =>
                 {
                     var val = prop.GetValue(x);


                     if (val == null) return null;

                     // Special handling for NodalSort
                     if (prop.Name.Equals("NodalSort", StringComparison.OrdinalIgnoreCase))
                     {
                         // If it�s numeric, fine � return it
                         if (double.TryParse(val.ToString(), out double nodalNum))
                             return nodalNum;

                         // ? Otherwise, throw to make the problem visible
                         throw new InvalidOperationException(
                             $"? NodalSort value is not numeric for record: {System.Text.Json.JsonSerializer.Serialize(x)} (actual value: '{val}')"
                         );
                     }

                     if (prop.Name.Equals("CenterSort", StringComparison.OrdinalIgnoreCase))
                     {
                         // If it�s numeric, fine � return it
                         if (int.TryParse(val.ToString(), out int centerNum))
                             return centerNum;

                         // ? Otherwise, throw to make the problem visible
                         throw new InvalidOperationException(
                             $"? CenterSort value is not numeric for record: {System.Text.Json.JsonSerializer.Serialize(x)} (actual value: '{val}')"
                         );
                     }

                     if (prop.Name.Equals("RouteSort", StringComparison.OrdinalIgnoreCase))
                     {
                         // If it�s numeric, fine � return it
                         if (int.TryParse(val.ToString(), out int routeNum))
                             return routeNum;

                         // ? Otherwise, throw to make the problem visible
                         throw new InvalidOperationException(
                             $"? RouteSort value is not numeric for record: {System.Text.Json.JsonSerializer.Serialize(x)} (actual value: '{val}')"
                         );
                     }
                     if (prop.Name.Equals("ExamDate", StringComparison.OrdinalIgnoreCase))
                     {
                         if (DateTime.TryParseExact(val.ToString(), "dd-MM-yyyy",
                             System.Globalization.CultureInfo.InvariantCulture,
                             System.Globalization.DateTimeStyles.None, out DateTime parsedDate))
                             return parsedDate;
                     }

                     if (val is DateTime dt)
                         return dt;

                     return val.ToString().Trim();

                 };

                 if (i == 0)
                     ordered = enrichedList.OrderBy(keySelector);
                 else
                     ordered = ordered.ThenBy(keySelector);
             }

             // Step 3: Update the list if sorting happened
             if (ordered != null)
                 enrichedList = ordered.ToList();

             var sortedList = ordered?.ToList() ?? enrichedList;

             // Step 6: Add TotalPages and BoxNo
             int boxNo = startBox;
             int runningPages = 0;
             string prevMergeKey = null;

             var finalWithBoxes = new List<dynamic>();
             long runningOmrPointer = 0;
             string previousCatchForOmr = null;
             string previousCourse = null;
             int innerBundlingSerial = 0;
             string prevInnerBundlingKey = null;
             foreach (var item in sortedList)
             {
                 bool hasOmr = !string.IsNullOrWhiteSpace(item.OmrSerial);
                 var nrRow = nrData.FirstOrDefault(n => n.CatchNo == item.CatchNo);
                     int pages = nrRow?.Pages ?? 0;
                 int totalPages = (item.Quantity) * pages;
                 string currentSymbol = resetOnSymbolChange
            ? (nrRow?.Symbol ?? "")
            : "";

                 string currentCourseName = item.CourseName?.ToString() ?? "";

                 // ? Reset boxNo when CourseName changes
                 if (resetOnSymbolChange && previousCourse != null && currentCourseName != previousCourse)
                 {
                     boxNo = startBox;
                     runningPages = 0;
                     _loggerService.LogEvent(
                         $"?? CourseName changed from '{previousCourse}' to '{currentCourseName}' ? BoxNo reset to {boxNo}",
                         "EnvelopeBreakages",
                         LogHelper.GetTriggeredBy(User),
                         ProjectId);
                 }

                 if (hasOmr)
                 {
                     if (previousCatchForOmr != item.CatchNo)
                     {
                         runningOmrPointer = 0;
                         previousCatchForOmr = item.CatchNo;
                     }
                     if (runningOmrPointer == 0 && item.OmrSerial.Contains("-"))
                     {
                         var parts = item.OmrSerial.Split('-');
                         runningOmrPointer = long.Parse(parts[0]);
                     }
                 }


                 string innerBundlingKey = null;
                 int currentInnerBundlingSerial = 0;
                 if (InnerBundling && innerBundlingFieldNames.Any())
                 {
                     innerBundlingKey = string.Join("_", innerBundlingFieldNames.Select(fieldName =>
                     {
                         var prop = item?.GetType().GetProperty(fieldName, BindingFlags.Public | BindingFlags.Instance | BindingFlags.IgnoreCase);
                         return prop?.GetValue(item)?.ToString()?.Trim() ?? "";
                     }));

                     if (innerBundlingKey != prevInnerBundlingKey)
                     {
                         innerBundlingSerial++;
                         prevInnerBundlingKey = innerBundlingKey;
                     }
                     currentInnerBundlingSerial = innerBundlingSerial;
                 }
                 // ? Helper: format BoxNo as concatenation of number + symbol
                 string FormatBoxNo(int num, string symbol)
                     => resetOnSymbolChange && !string.IsNullOrEmpty(symbol)
                         ? $"{num}{symbol}"
                         : num.ToString();
                 // Build merge key
                 string mergeKey = "";
                 if (boxIds == null)
                 {
                     _loggerService.LogError("Capacity is not found", "", nameof(EnvelopeBreakagesController));
                     return NotFound("Capacity is not found");
                 }
                 if (boxIds.Any())
                 {
                     mergeKey = string.Join("_", boxIds.Select(fieldId =>
                     {
                         var fieldName = fields.FirstOrDefault(f => f.FieldId == fieldId)?.Name;
                         if (fieldName != null)
                         {
                             var prop = item?.GetType().GetProperty(fieldName, BindingFlags.Public | BindingFlags.Instance | BindingFlags.IgnoreCase);

                             if (prop != null)
                             {
                                 var value = prop.GetValue(item)?.ToString() ?? "";
                                 return value;
                             }
                             else
                             {
                                 _loggerService.LogError($"Property '{fieldName}' not found on nrRow", "", nameof(EnvelopeBreakagesController));

                             }
                         }
                         else
                         {
                             _loggerService.LogError($"Field name not found for fieldId: {fieldId}", "", nameof(EnvelopeBreakagesController));

                         }

                         return ""; // fallback if field name or property not found
                     }));
                 }
                 // ---- Rule 1: merge fields change ? force new box
                 bool mergeChanged = (prevMergeKey != null && mergeKey != prevMergeKey);

                 try
                 {
                     if (mergeChanged)
                     {
                         boxNo++; // start new box for new merge group
                         runningPages = 0;
                         _loggerService.LogEvent($"?? MergeKey changed ? new box {boxNo} for {mergeKey}",
                             "EnvelopeBreakages",
                             LogHelper.GetTriggeredBy(User),
                             ProjectId);
                     }
                     bool overflow = (runningPages + totalPages > capacity);
                     if (overflow)
                     {
                         _loggerService.LogEvent($"Overflow {string.Join(", ", mergeKey)} Running {runningPages} Total Pages {totalPages}, Capacity {capacity} boxNo {boxNo}" +
                                $"", "EnvelopeBreakages", LogHelper.GetTriggeredBy(User), ProjectId);
                         Console.WriteLine(envelopeSize);
                         if (envelopeSize < 50 && envelopeSize > 0)
                         {
                             int pagesPerUnit = (nrRow?.Pages ?? 0);
                             int remainingQty = item.Quantity;

                             var lastBoxForCatch = finalWithBoxes
                                 .Where(b => (string)b.GetType().GetProperty("CatchNo").GetValue(b) == item.CatchNo)
                                 .OrderBy(b => (int)b.GetType().GetProperty("End").GetValue(b))
                                 .LastOrDefault();

                             int currentStart = lastBoxForCatch != null
                                 ? (int)lastBoxForCatch.GetType().GetProperty("End").GetValue(lastBoxForCatch) + 1
                                 : (int)item.Start;

                             while (remainingQty > 0)
                             {
                                 // how many pages can still fit in current box
                                 int remainingCapacity = capacity - runningPages;

                                 // max qty that fits in remaining capacity, rounded down to envelopeSize
                                 int maxFittingQty = (int)Math.Floor((double)remainingCapacity / pagesPerUnit);
                                 maxFittingQty = (maxFittingQty / envelopeSize) * envelopeSize;
                                 if (maxFittingQty <= 0)
                                 {
                                     // current box is full, open new box
                                     boxNo++;
                                     runningPages = 0;
                                     if (InnerBundling)
                                     {
                                         innerBundlingSerial++;
                                         currentInnerBundlingSerial = innerBundlingSerial;
                                     }
                                     continue;
                                 }

                                 int chunkQty = Math.Min(maxFittingQty, remainingQty);
                                 // round down to envelopeSize
                                 chunkQty = (chunkQty / envelopeSize) * envelopeSize;
                                 if (chunkQty <= 0)
                                 {
                                     boxNo++;
                                     runningPages = 0;
                                     if (InnerBundling)
                                     {
                                         innerBundlingSerial++;
                                         currentInnerBundlingSerial = innerBundlingSerial;
                                     }
                                     continue;
                                 }

                                 int chunkPages = chunkQty * pagesPerUnit;
                                 int envelopesInBox = chunkQty / envelopeSize;
                                 int start = currentStart;
                                 int end = start + envelopesInBox - 1;
                                 currentStart = end + 1;
                                 string serial = $"{start} to {end}";
                                 string omrRange = "";

                                 if (hasOmr)
                                 {
                                     long omrStart = runningOmrPointer;
                                     long omrEnd = omrStart + chunkQty - 1;
                                     omrRange = $"{omrStart}-{omrEnd}";
                                     runningOmrPointer = omrEnd + 1;
                                 }
                                 object boxNoValue = resetOnSymbolChange
                            ? (object)$"{boxNo}{currentSymbol}"
                            : boxNo;
                                 if (InnerBundling)
                                 {
                                     finalWithBoxes.Add(new
                                     {
                                         item.CatchNo,
                                         item.CenterCode,
                                         item.CenterSort,
                                         item.ExamTime,
                                         item.ExamDate,
                                         Quantity = chunkQty,
                                         item.NodalCode,
                                         item.NodalSort,
                                         item.Route,
                                         item.RouteSort,
                                         item.TotalEnv,
                                         Start = start,
                                         End = end,
                                         Serial = serial,
                                         TotalPages = chunkPages,
                                         BoxNo = boxNoValue,
                                         OmrSerial = omrRange,
                                         InnerBundlingSerial = currentInnerBundlingSerial,
                                         item.CourseName,
                                     });
                                 }
                                 else
                                 {
                                     finalWithBoxes.Add(new
                                     {
                                         item.CatchNo,
                                         item.CenterCode,
                                         item.CenterSort,
                                         item.ExamTime,
                                         item.ExamDate,
                                         Quantity = chunkQty,
                                         item.NodalCode,
                                         item.NodalSort,
                                         item.Route,
                                         item.RouteSort,
                                         item.TotalEnv,
                                         Start = start,
                                         End = end,
                                         Serial = serial,
                                         TotalPages = chunkPages,
                                         BoxNo = boxNoValue,
                                         OmrSerial = omrRange,
                                         item.CourseName,
                                     });
                                 }

                                     runningPages += chunkPages; // ? naturally tracks current box pages
                                 remainingQty -= chunkQty;
                             }

                             prevMergeKey = mergeKey;
                             previousCourse = currentCourseName;
                             continue;   
                         }
                         boxNo++;
                         runningPages = 0;
                     }


                         // normal case: just start new box
                     runningPages += totalPages;
                     string normalOmrRange = "";

                     if (hasOmr)
                     {
                         long normalOmrStart = runningOmrPointer;
                         long normalOmrEnd = normalOmrStart + item.Quantity - 1;
                         normalOmrRange = $"{normalOmrStart}-{normalOmrEnd}";
                         runningOmrPointer = normalOmrEnd + 1;
                     }
                     object normalBoxNoValue = resetOnSymbolChange
                ? (object)$"{currentSymbol}-{boxNo}"
                : boxNo;
                     if (InnerBundling)
                     {
                         finalWithBoxes.Add(new
                         {
                             item.CatchNo,
                             item.CenterCode,
                             item.CenterSort,
                             item.ExamTime,
                             item.ExamDate,
                             item.Quantity,
                             item.NodalCode,
                             item.NodalSort,
                             item.Route,
                             item.RouteSort,
                             item.TotalEnv,
                             item.Start,
                             item.End,
                             item.Serial,
                             TotalPages = totalPages,
                             BoxNo = normalBoxNoValue,
                             OmrSerial = normalOmrRange,
                             InnerBundlingSerial = currentInnerBundlingSerial,
                             item.CourseName
                         });
                     }
                     else
                     {
                         finalWithBoxes.Add(new
                         {
                             item.CatchNo,
                             item.CenterCode,
                             item.CenterSort,
                             item.ExamTime,
                             item.ExamDate,
                             item.Quantity,
                             item.NodalCode,
                             item.NodalSort,
                             item.Route,
                             item.RouteSort,
                             item.TotalEnv,
                             item.Start,
                             item.End,
                             item.Serial,
                             TotalPages = totalPages,
                             BoxNo = normalBoxNoValue,
                             OmrSerial = normalOmrRange,
                             item.CourseName,
                         });
                     }

                     prevMergeKey = mergeKey;
                     previousCourse = currentCourseName;
                 }

                 catch (Exception ex)
                 {
                     _loggerService.LogError("Error storing report EnvelopeBreakage", ex.Message, nameof(EnvelopeBreakagesController));
                     return StatusCode(500, "Internal Server Error");
                 }
             }

             // ?? Maintain ordering
             *//*  finalWithBoxes = resetOnSymbolChange
                   ? finalWithBoxes
                       .OrderBy(x => x.CourseName?.ToString() ?? "")  // ? group by course first
                       .ThenBy(x =>
                       {
                           string boxNoStr = x.BoxNo?.ToString() ?? "";
                           string numPart = new string(boxNoStr.TakeWhile(char.IsDigit).ToArray());
                           return int.TryParse(numPart, out int n) ? n : 0;
                       })
                       .ToList()
                   : finalWithBoxes.OrderBy(x => (int)x.BoxNo).ToList();
   *//*

             // Step 5: Export to Excel
             try
             {
                 using (var package = new ExcelPackage())
                 {
                     var worksheet = package.Workbook.Worksheets.Add("BoxBreaking");

                     // ? Detect OMR column
                     bool hasAnyOmr = finalWithBoxes.Any(x =>
                         !string.IsNullOrWhiteSpace(
                             x.GetType().GetProperty("OmrSerial")?.GetValue(x)?.ToString()));
                     var fixedColumns = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
 {
     "CatchNo",
     "CenterCode",
     "CenterSort",
     "ExamTime",
     "ExamDate",
     "Quantity",
     "NodalCode",
     "NodalSort",
     "Route",
     "RouteSort",
     "TotalEnv",
     "Start",
     "End",
     "Serial",
     "Pages",
     "TotalPages",
     "BoxNo",
     "OmrSerial",
     "CourseName",
     "Symbol",
     "InnerBundlingSerial"
 };
                     // ? Dynamically get extra NR columns
                     var extraNRColumns = typeof(NRData)
     .GetProperties()
     .Where(p => p.Name != "NRDatas" && p.Name != "Id" && p.Name != "ProjectId" && !fixedColumns.Contains(p.Name))
     .Where(p => nrData.Any(n => p.GetValue(n) != null)) // only if data exists
     .Select(p => p.Name)
     .ToList();

                     // ? Extract JSON keys dynamically
                     List<string> jsonKeys = new List<string>();

                     if (nrData.Any(n => !string.IsNullOrEmpty(n.NRDatas)))
                     {
                         var sampleJson = nrData.First(n => !string.IsNullOrEmpty(n.NRDatas)).NRDatas;

                         var dict = JsonSerializer.Deserialize<Dictionary<string, object>>(sampleJson);
                         jsonKeys = dict?.Keys.ToList() ?? new List<string>();
                     }

                     // ================= HEADER =================
                     worksheet.Cells[1, 1].Value = "SerialNumber";
                     worksheet.Cells[1, 2].Value = "CatchNo";
                     worksheet.Cells[1, 3].Value = "CenterCode";
                     worksheet.Cells[1, 4].Value = "CenterSort";
                     worksheet.Cells[1, 5].Value = "ExamTime";
                     worksheet.Cells[1, 6].Value = "ExamDate";
                     worksheet.Cells[1, 7].Value = "Quantity";
                     worksheet.Cells[1, 8].Value = "NodalCode";
                     worksheet.Cells[1, 9].Value = "NodalSort";
                     worksheet.Cells[1, 10].Value = "Route";
                     worksheet.Cells[1, 11].Value = "RouteSort";
                     worksheet.Cells[1, 12].Value = "TotalEnv";
                     worksheet.Cells[1, 13].Value = "Start";
                     worksheet.Cells[1, 14].Value = "End";
                     worksheet.Cells[1, 15].Value = "Serial";
                     worksheet.Cells[1, 16].Value = "Pages";
                     worksheet.Cells[1, 17].Value = "TotalPages";

                     int nextCol = 18;
                     int symbolCol = -1;
                     int courseCol = -1;

                     if (resetOnSymbolChange)
                     {
                         worksheet.Cells[1, nextCol].Value = "Symbol";
                         symbolCol = nextCol++;

                         worksheet.Cells[1, nextCol].Value = "CourseName";
                         courseCol = nextCol++;
                     }

                     worksheet.Cells[1, nextCol].Value = "BoxNo";
                     int boxCol = nextCol++;

                     worksheet.Cells[1, nextCol].Value = "Beejak";
                     int beejakCol = nextCol++;

                     int omrCol = -1;
                     if (hasAnyOmr)
                     {
                         worksheet.Cells[1, nextCol].Value = "OmrSerial";
                         omrCol = nextCol++;
                     }

                     int innerBundlingCol = -1;
                     if (InnerBundling)
                     {
                         worksheet.Cells[1, nextCol].Value = "InnerBundlingSerial";
                         innerBundlingCol = nextCol++;
                     }

                     // ? Extra NR Headers
                     int extraStartCol = nextCol;

                     foreach (var prop in extraNRColumns)
                     {
                         worksheet.Cells[1, nextCol].Value = prop;
                         nextCol++;
                     }

                     // ? JSON Headers
                     foreach (var key in jsonKeys)
                     {
                         worksheet.Cells[1, nextCol].Value = key;
                         nextCol++;
                     }

                     // ================= DATA =================
                     int row = 2;
                     int serial = 1;
                     string previousCenter = null;
                     string previousBox = null;

                     foreach (var item in finalWithBoxes)
                     {
                         var nrRow = nrData.FirstOrDefault(n => n.CatchNo == item.CatchNo);

                         worksheet.Cells[row, 1].Value = serial++;
                         worksheet.Cells[row, 2].Value = item.CatchNo;
                         worksheet.Cells[row, 3].Value = item.CenterCode;
                         worksheet.Cells[row, 4].Value = item.CenterSort;
                         worksheet.Cells[row, 5].Value = item.ExamTime;
                         worksheet.Cells[row, 6].Value = item.ExamDate;
                         worksheet.Cells[row, 7].Value = item.Quantity;
                         worksheet.Cells[row, 8].Value = item.NodalCode;
                         worksheet.Cells[row, 9].Value = item.NodalSort;
                         worksheet.Cells[row, 10].Value = item.Route;
                         worksheet.Cells[row, 11].Value = item.RouteSort;
                         worksheet.Cells[row, 12].Value = item.TotalEnv;
                         worksheet.Cells[row, 13].Value = item.Start;
                         worksheet.Cells[row, 14].Value = item.End;
                         worksheet.Cells[row, 15].Value = item.Serial;
                         worksheet.Cells[row, 16].Value = nrRow?.Pages ?? 0;
                         worksheet.Cells[row, 17].Value = item.TotalPages;

                         worksheet.Cells[row, boxCol].Value = item.BoxNo;

                         if (omrCol > 0)
                             worksheet.Cells[row, omrCol].Value = item.OmrSerial;

                         if (symbolCol > 0)
                             worksheet.Cells[row, symbolCol].Value = nrRow?.Symbol ?? "";

                         if (courseCol > 0)
                             worksheet.Cells[row, courseCol].Value = item.CourseName;

                         // ?? Beejak logic
                         string currentCenter = item.CenterCode?.ToString();
                         string currentBox = item.BoxNo?.ToString();

                         string beejakValue = "";
                         if (currentBox != previousBox || currentCenter != previousCenter)
                             beejakValue = "Beejak";

                         worksheet.Cells[row, beejakCol].Value = beejakValue;

                         previousCenter = currentCenter;
                         previousBox = currentBox;

                         if (innerBundlingCol > 0)
                             worksheet.Cells[row, innerBundlingCol].Value = item.InnerBundlingSerial;

                         // ? Dynamic columns
                         int currentCol = extraStartCol;

                         foreach (var prop in extraNRColumns)
                         {
                             var val = nrRow?.GetType().GetProperty(prop)?.GetValue(nrRow);
                             worksheet.Cells[row, currentCol++].Value = val;
                         }

                         Dictionary<string, object> jsonDict = null;

                         if (!string.IsNullOrEmpty(nrRow?.NRDatas))
                         {
                             jsonDict = JsonSerializer.Deserialize<Dictionary<string, object>>(nrRow.NRDatas);
                         }

                         foreach (var key in jsonKeys)
                         {
                             worksheet.Cells[row, currentCol++].Value =
                                 (jsonDict != null && jsonDict.ContainsKey(key))
                                 ? jsonDict[key]?.ToString()
                                 : "";
                         }

                         row++;
                     }

                     FileInfo fi = new FileInfo(filePath);
                     package.SaveAs(fi);
                 }

                 // ? API call
                 using var client = new HttpClient();
                 var response = await client.PostAsync(
                     $"{_apiSettings.BoxBreaking}?ProjectId={ProjectId}",
                     new StringContent("")
                 );

                 if (!response.IsSuccessStatusCode)
                 {
                     var error = await response.Content.ReadAsStringAsync();
                     Console.WriteLine($"API Failed: {response.StatusCode}, {error}");
                 }
                 else
                 {
                     var data = await response.Content.ReadAsStringAsync();
                     Console.WriteLine($"API Success: {data}");
                 }

                 return Ok(new { message = "File successfully created", filePath });
             }
             catch (Exception ex)
             {
                 _loggerService.LogError(
                     "Error creating report BoxBreakingReport",
                     ex.Message,
                     nameof(EnvelopeBreakagesController)
                 );

                 return StatusCode(500, "Internal Server Error");
             }
         }

         public class ExcelInputRow
         {
             public string CatchNo { get; set; }
             public string CenterCode { get; set; }
             public string ExamTime { get; set; }
             public string ExamDate { get; set; }
             public int Quantity { get; set; }
             public int TotalEnv { get; set; }
             public int NRQuantity { get; set; }
             public string NodalCode { get; set; }
             public int CenterSort { get; set; }
             public double NodalSort { get; set; }
             public string Route {  get; set; }
             public int RouteSort { get; set; }
             public string OmrSerial { get; set; }
             public string CourseName { get; set; }
         }
 */

        /*  [HttpGet("EnvelopeBreakage")]
          public async Task<IActionResult> BreakageConfiguration(int ProjectId)
          {
              // Fetch all data sequentially

              var envCaps = await _context.EnvelopesTypes
                  .Select(e => new { e.EnvelopeName, e.Capacity })
                  .ToListAsync();

              var nrData = await _context.NRDatas
                  .Where(p => p.ProjectId == ProjectId && p.Status == true)
                  .OrderBy(p=>p.CatchNo)
                  .ThenBy(p => p.RouteSort)
                  .ThenBy(p => p.NodalSort)
                  .ThenBy(p => p.CenterSort)
                  .ToListAsync();


              var envBreaking = await _context.EnvelopeBreakages
                  .Where(p => p.ProjectId == ProjectId)
                  .ToListAsync();

              var extras = await _context.ExtrasEnvelope
                  .Where(p => p.ProjectId == ProjectId && p.Status == 1)
                  .ToListAsync();

              var projectconfig = await _context.ProjectConfigs
                  .Where(p => p.ProjectId == ProjectId)
                  .FirstOrDefaultAsync();

              var outerEnvJson = projectconfig.Envelope;
              bool resetOmrSerialOnCatchChange = projectconfig.ResetOmrSerialOnCatchChange;
              var startNumber = projectconfig.OmrSerialNumber;

              var EnvelopeBreaking = projectconfig.EnvelopeMakingCriteria;
              var fields = await _context.Fields.Where(f => EnvelopeBreaking.Contains(f.FieldId)).ToListAsync();
              var fieldNames = fields.OrderBy(f => EnvelopeBreaking.IndexOf(f.FieldId)).Select(f => f.Name); // Get the field names .ToList();

              var extrasconfig = await _context.ExtraConfigurations
                  .Where(p => p.ProjectId == ProjectId)
                  .ToListAsync();

              // Build dictionaries for fast lookup
              var envelopeCapacities = envCaps.ToDictionary(x => x.EnvelopeName, x => x.Capacity);
              var envDict = envBreaking.ToDictionary(
                 e => e.NrDataId,
                 e => (e.TotalEnvelope, e.OuterEnvelope) // ?? this is a tuple
              );

              var mssTypes = projectconfig.MssTypes; // assuming List<int>
              var mssAttached = projectconfig.MssAttached;
              string mssMode = projectconfig.MssAttached?.ToLower();
              var mssData = await _context.Mss
                  .Where(m => mssTypes.Contains(m.Id))
                  .ToListAsync();
              var resultList = new List<object>();
              string prevNodalCode = null;
              string prevRoute = null;
              int prevRouteSort = 0;
              double prevNodalSort = 0;
              string prevCatchNo = null;
              string prevMergeField = null;
              string prevExtraMergeField = null;
              int centerEnvCounter = 0;
              int extraCenterEnvCounter = 0;
              var nodalExtrasAddedForNodalCatch = new HashSet<(string NodalCode, string CatchNo)>();
              var catchExtrasAdded = new HashSet<(int ExtraId, string CatchNo)>();
              List<object> CreateMssRows(string catchNo, string examDate, string examTime, string courseName)
              {
                  var rows = new List<object>();

                  foreach (var mss in mssData)
                  {
                      rows.Add(new
                      {
                          CatchNo = catchNo,
                          CenterCode = "",
                          CourseName = courseName,   // ? added
                          ExamTime = examTime,       // ? added
                          ExamDate = examDate,       // ? added
                          Quantity = "",
                          EnvQuantity = mss.MssType,    // MSS Type Name
                          NodalCode = "",
                          CenterEnv = "",
                          TotalEnv = "",
                          Env = "",
                          NRQuantity = "",
                          CenterSort = "",
                          NodalSort = "",
                          Route = "",
                          RouteSort = "",
                          isMss = true,
                      });
                  }

                  return rows;
              }
              // Helper method to add extra envelopes - removed serialnumber++ from here
              void AddExtraWithEnv(ExtraEnvelopes extra, string examDate, string examTime, string course,int NrQuantity, string NodalCode, string CenterCode, double CenterSort, double NodalSort, int RouteSort, string Route)
              {
                  var extraConfig = extrasconfig.FirstOrDefault(e => e.ExtraType == extra.ExtraId);
                  int envCapacity = 0; // default fallback

                  if (extraConfig != null && !string.IsNullOrEmpty(extraConfig.EnvelopeType))
                  {
                      var envType = JsonSerializer.Deserialize<Dictionary<string, string>>(extraConfig.EnvelopeType);
                      if (envType != null && envType.TryGetValue("Outer", out string outerType)
                          && envelopeCapacities.TryGetValue(outerType, out int cap))
                      {
                          envCapacity = cap;
                      }
                  }

                  int totalEnv = (int)Math.Ceiling((double)extra.Quantity / envCapacity);
                  string currentMergeField = $"{extra.CatchNo}-{extra.ExtraId}";

                  if (currentMergeField != prevExtraMergeField)
                  {
                      extraCenterEnvCounter = 0;
                      prevExtraMergeField = currentMergeField;
                  }

                  for (int j = 1; j <= totalEnv; j++)
                  {
                      int envQuantity;
                      if (j == 1 && totalEnv > 1)
                      {
                          envQuantity = extra.Quantity - (envCapacity * (totalEnv - 1));
                      }
                      else
                      {
                          envQuantity = Math.Min(extra.Quantity - (envCapacity * (j - 1)), envCapacity);
                      }
                      extraCenterEnvCounter++;

                      resultList.Add(new
                      {
                          ExtraAttached = true,
                          extra.ExtraId,
                          extra.CatchNo,
                          extra.Quantity,
                          EnvQuantity = envQuantity,
                          extra.InnerEnvelope,
                          extra.OuterEnvelope,
                          CenterCode = extra.ExtraId switch
                          {
                              1 => "Nodal Extra",
                              2 => "University Extra",
                              3 => "Office Extra",
                              _ => "Extra"
                          },
                          CenterEnv = extraCenterEnvCounter,
                          ExamDate = examDate,
                          CourseName = course,
                          ExamTime = examTime,
                          TotalEnv = totalEnv,
                          Env = $"{j}/{totalEnv}",
                          NRQuantity = NrQuantity,
                          NodalCode = NodalCode,
                          Route = "",
                          NodalSort = extra.ExtraId switch
                          {
                              1 =>(int)NodalSort + 0.1,
                              2 => 100000,
                              3 =>1000000,
                          },
                            CenterSort = extra.ExtraId switch
                            {
                                1 => 10000,
                                2 => 100000,
                                3 => 1000000,
                            },
                            RouteSort = extra.ExtraId switch
                            {
                                1 => (int)RouteSort,
                                2 => 10000,
                                3 => 100000,
                            },
                      });
                  }
              }

              for (int i = 0; i < nrData.Count; i++)
              {
                  var current = nrData[i];
                  Console.WriteLine($" Processing: {nrData[i].CatchNo}, CenterSort={nrData[i].CenterSort}, NodalSort={nrData[i].NodalSort}");
                  // Check if CatchNo changed BEFORE processing extras
                  bool catchNoChanged = prevCatchNo != null && current.CatchNo != prevCatchNo;

                  if (catchNoChanged)
                  {
                      var prevNrData = nrData[i - 1];
                      // ? Add final extras for previous CatchNo before resetting serial
                      if (!nodalExtrasAddedForNodalCatch.Contains((prevNrData.NodalCode, prevCatchNo)))
                      {
                          var extrasToAdd = extras.Where(e => e.ExtraId == 1 && e.CatchNo == prevCatchNo).ToList();
                          // Get previous record for metadata
                          foreach (var extra in extrasToAdd)
                          {
                              AddExtraWithEnv(extra, prevNrData.ExamDate, prevNrData.ExamTime, prevNrData.CourseName,
                                            prevNrData.NRQuantity, prevNrData.NodalCode, prevNrData.CenterCode, prevNrData.CenterSort, prevNrData.NodalSort,prevNrData.RouteSort, prevNrData.Route);
                          }
                          nodalExtrasAddedForNodalCatch.Add((prevNrData.NodalCode, prevCatchNo));
                      }

                      foreach (var extraId in new[] { 2, 3 })
                      {
                          if (!catchExtrasAdded.Contains((extraId, prevCatchNo)))
                          {
                              var extrasToAdd = extras.Where(e => e.ExtraId == extraId && e.CatchNo == prevCatchNo).ToList();
                              foreach (var extra in extrasToAdd)
                              {
                                  AddExtraWithEnv(extra, prevNrData.ExamDate, prevNrData.ExamTime,prevNrData.CourseName,
                                                 prevNrData.NRQuantity, prevNrData.NodalCode, prevNrData.CenterCode, prevNrData.CenterSort, prevNrData.NodalSort,prevNrData.RouteSort,prevNrData.Route);
                              }
                              catchExtrasAdded.Add((extraId, prevCatchNo));
                          }
                      }

                      // NOW reset serial number after extras are added
                  }

                  // ? Nodal Extra when NodalCode changes (but not CatchNo)
                  if (!catchNoChanged && prevNodalCode != null && current.NodalCode != prevNodalCode)
                  {
                      Console.WriteLine($" Processing: {nrData[i].CatchNo}, CenterSort={nrData[i].CenterSort}, NodalSort={nrData[i].NodalSort}");

                      if (!nodalExtrasAddedForNodalCatch.Contains((prevNodalCode, current.CatchNo)))
                      {
                          var extrasToAdd = extras.Where(e => e.ExtraId == 1 && e.CatchNo == current.CatchNo).ToList();
                          foreach (var extra in extrasToAdd)
                          {
                              AddExtraWithEnv(extra, current.ExamDate, current.ExamTime,current.CourseName,
                                             current.NRQuantity, prevNodalCode, current.CenterCode, current.CenterSort, prevNodalSort, prevRouteSort, prevRoute);
                          }
                          nodalExtrasAddedForNodalCatch.Add((prevNodalCode, current.CatchNo));
                      }
                  }

                  // ? Add current NRData row with TotalEnv replication
                  int totalEnv = envDict.TryGetValue(current.Id, out var envData) && envData.TotalEnvelope > 0
                  ? envData.TotalEnvelope
                 : 1;

                  // Calculate CenterEnv
                  string currentMergeField = $"{current.CatchNo}-{current.CenterCode}";
                  if (currentMergeField != prevMergeField)
                  {
                      centerEnvCounter = 0;
                      prevMergeField = currentMergeField;
                  }

                  var envelopeBreakdown = new List<(string EnvType, int Count, int Capacity)>();

                  if (envDict.TryGetValue(current.Id, out var envInfo) && !string.IsNullOrEmpty(envInfo.OuterEnvelope))
                  {
                      try
                      {
                          var outerEnvDict = JsonSerializer.Deserialize<Dictionary<string, string>>(envInfo.OuterEnvelope);
                          if (outerEnvDict != null)
                          {
                              foreach (var kvp in outerEnvDict)
                              {
                                  if (int.TryParse(kvp.Value, out int count) && count > 0)
                                  {
                                      // Get capacity from envelope name (e.g., "E50" -> 50)
                                      int capacity = 0;
                                      if (envelopeCapacities.TryGetValue(kvp.Key, out int cap))
                                      {
                                          capacity = cap;
                                      }
                                      else
                                      {
                                          // Try to parse capacity from envelope name (e.g., E50 -> 50)
                                          var match = System.Text.RegularExpressions.Regex.Match(kvp.Key, @"\d+");
                                          if (match.Success)
                                          {
                                              capacity = int.Parse(match.Value);
                                          }
                                      }

                                      envelopeBreakdown.Add((kvp.Key, count, capacity));
                                  }
                              }
                          }
                      }
                      catch { }
                  }

                  // Sort by capacity (ascending - smallest first)
                  envelopeBreakdown = envelopeBreakdown.OrderBy(x => x.Capacity).ToList();
                  // If no breakdown found, use default behavior

                  if (envelopeBreakdown.Count > 0)
                  {
                      // Process each envelope type from the breakdown
                      int envelopeIndex = 1;
                      int remainingQty = current.Quantity;
                      foreach (var (envType, count, capacity) in envelopeBreakdown)
                      {
                          for (int k = 0; k < count; k++)
                          {
                              centerEnvCounter++;
                              int envQty;
                              if (remainingQty > capacity)
                              {
                                  envQty = capacity;
                              }
                              else
                              {
                                  envQty = remainingQty;

                              }
                              resultList.Add(new
                              {
                                  current.CatchNo,
                                  current.CenterCode,
                                  current.CourseName,
                                  current.ExamTime,
                                  current.ExamDate,
                                  current.Quantity,
                                  EnvQuantity = envQty,  // Use the envelope capacity
                                  current.NodalCode,
                                  CenterEnv = centerEnvCounter,
                                  TotalEnv = totalEnv,
                                  Env = $"{envelopeIndex}/{totalEnv}",
                                  current.NRQuantity,
                                  current.CenterSort,
                                  current.NodalSort,
                                  current.Route,
                                  current.RouteSort,
                                  isMss = false,
                              });
                              remainingQty -= envQty;
                              envelopeIndex++;
                              if (remainingQty <= 0)
                                  break;
                          }
                          if (remainingQty <= 0)
                              break;
                      }
                  }

                  prevNodalCode = current.NodalCode;
                  prevNodalSort = current.NodalSort;
                  prevRouteSort = current.RouteSort;
                  prevRoute = current.Route;
                  prevCatchNo = current.CatchNo;
              }

              // ?? Final extras for the last CatchNo
              if (prevCatchNo != null)
              {
                  var lastNrData = nrData.LastOrDefault();

                  if (lastNrData != null)
                  {
                      if (!nodalExtrasAddedForNodalCatch.Contains((lastNrData.NodalCode, prevCatchNo)))
                      {
                          var extrasToAdd = extras.Where(e => e.ExtraId == 1 && e.CatchNo == prevCatchNo).ToList();
                          foreach (var extra in extrasToAdd)
                          {
                              AddExtraWithEnv(extra, lastNrData.ExamDate, lastNrData.ExamTime,lastNrData.CourseName,
                                         lastNrData.NRQuantity, lastNrData.NodalCode, lastNrData.CenterCode, lastNrData.CenterSort, lastNrData.NodalSort,lastNrData.RouteSort, lastNrData.Route);
                          }
                      }

                      foreach (var extraId in new[] { 2, 3 })
                      {
                          if (!catchExtrasAdded.Contains((extraId, prevCatchNo)))
                          {
                              var extrasToAdd = extras.Where(e => e.ExtraId == extraId && e.CatchNo == prevCatchNo).ToList();
                              foreach (var extra in extrasToAdd)
                              {
                                  AddExtraWithEnv(extra, lastNrData.ExamDate, lastNrData.ExamTime,lastNrData.CourseName,
                                               lastNrData.NRQuantity, lastNrData.NodalCode, lastNrData.CenterCode,lastNrData.CenterSort,lastNrData.NodalSort,lastNrData.RouteSort, lastNrData.Route);
                              }
                          }
                      }
                  }
              }
              int currentStartNumber = startNumber;
              bool assignBookletSerial = currentStartNumber > 0;
              string prevCatchForSerial = null;
              var nonMssRows = resultList
      .Where(r => {
          var p = r.GetType().GetProperty("isMss");
          return p == null || !(p.GetValue(r) is bool b && b);
      }).ToList();
              // Generate Excel Report
              // Sort the resultList safely (string for CatchNo, numeric for CenterSort/NodalSort)
              IOrderedEnumerable<dynamic> ordered = null;

              foreach (var fieldName in fieldNames)
              {
                  Func<dynamic, object> keySelector = record =>
                  {
                      var prop = record.GetType().GetProperty(fieldName);
                      if (prop == null) return null;

                      var val = prop.GetValue(record);
                      if (val == null) return null;

                      // ---- TYPE-SAFE HANDLING PER FIELD ----
                      switch (fieldName)
                      {
                          case "RouteSort":
                              if (int.TryParse(val.ToString(), out int intVal))
                                  return intVal;
                              return 0;
                          case "CenterSort":
                              if (int.TryParse(val.ToString(), out int intval))
                                  return intval;
                              return 0;

                          case "NodalSort":
                              if (double.TryParse(val.ToString(), out double dblVal))
                                  return dblVal;
                              return 0.0;

                          default:
                              return val.ToString().Trim();
                      }
                  };

                  if (ordered == null)
                      ordered = nonMssRows.Cast<dynamic>().OrderBy(keySelector);
                  else
                      ordered = ordered.ThenBy(keySelector);
              }

              var sortedNonMss = ordered?.Cast<object>().ToList() ?? nonMssRows;
              var finalList = new List<object>();
              string lastCatchForMss = null;
              var buffer = new List<object>();

              foreach (var item in sortedNonMss)
              {
                  string catchNo = item.GetType().GetProperty("CatchNo")?.GetValue(item)?.ToString();

                  if (catchNo != lastCatchForMss && lastCatchForMss != null)
                  {
                      if (mssMode == "end")
                      {
                          // flush buffer then add MSS after previous catch
                          finalList.AddRange(buffer);
                          var last = buffer.Last();
                          finalList.AddRange(CreateMssRows(
                              lastCatchForMss,
                              last.GetType().GetProperty("ExamDate")?.GetValue(last)?.ToString(),
                              last.GetType().GetProperty("ExamTime")?.GetValue(last)?.ToString(),
                              last.GetType().GetProperty("CourseName")?.GetValue(last)?.ToString()
                          ));
                      }
                      else if (mssMode == "start")
                      {
                          // flush buffer then add MSS before next catch
                          finalList.AddRange(buffer);
                          finalList.AddRange(CreateMssRows(
                              catchNo,
                              item.GetType().GetProperty("ExamDate")?.GetValue(item)?.ToString(),
                              item.GetType().GetProperty("ExamTime")?.GetValue(item)?.ToString(),
                              item.GetType().GetProperty("CourseName")?.GetValue(item)?.ToString()
                          ));
                      }
                      buffer.Clear();
                  }

                  // For the very first catch in start mode, add MSS before first item
                  if (mssMode == "start" && lastCatchForMss == null)
                  {
                      finalList.AddRange(CreateMssRows(
                          catchNo,
                          item.GetType().GetProperty("ExamDate")?.GetValue(item)?.ToString(),
                          item.GetType().GetProperty("ExamTime")?.GetValue(item)?.ToString(),
                          item.GetType().GetProperty("CourseName")?.GetValue(item)?.ToString()
                      ));
                  }

                  buffer.Add(item);
                  lastCatchForMss = catchNo;
              }

              // Flush last buffer
              if (buffer.Count > 0)
              {
                  finalList.AddRange(buffer);
                  if (mssMode == "end")
                  {
                      var last = buffer.Last();
                      finalList.AddRange(CreateMssRows(
                          lastCatchForMss,
                          last.GetType().GetProperty("ExamDate")?.GetValue(last)?.ToString(),
                          last.GetType().GetProperty("ExamTime")?.GetValue(last)?.ToString(),
                          last.GetType().GetProperty("CourseName")?.GetValue(last)?.ToString()
                      ));
                  }
              }

              resultList = finalList;

              // Generate SerialNumber AFTER sorting and reset per CatchNo
              // Generate SerialNumber AFTER sorting and reset per CatchNo
              int serial = 1;
              string previousCatchNo = null;

              var updatedList = new List<object>();

              foreach (var item in resultList)
              {
                  var type = item.GetType();
                  var props = type.GetProperties();
                  var isMssProp = props.FirstOrDefault(p => p.Name == "isMss");
                  bool isMssRow = isMssProp != null && isMssProp.GetValue(item) is bool b && b;

                  string currentCatchNo = props
                      .FirstOrDefault(p => p.Name == "CatchNo")
                      ?.GetValue(item)?.ToString();

                  if (!isMssRow && previousCatchNo != null && currentCatchNo != previousCatchNo)
                  {
                      serial = 1; // reset when CatchNo changes
                  }

                  var dict = new Dictionary<string, object>();

                  foreach (var prop in props)
                  {
                      dict[prop.Name] = prop.GetValue(item);
                  }

                  dict["SerialNumber"] = isMssRow ? (object)"" : serial++;

                  updatedList.Add(dict);
                  if (!isMssRow)
                      previousCatchNo = currentCatchNo;
              }

              resultList = updatedList;

              using (var package = new ExcelPackage())
              {
                  var worksheet = package.Workbook.Worksheets.Add("BreakingResult");

                  // Add headers
                  var headers = new[] { "Serial Number", "Catch No", "Center Code",
                          "Center Sort", "Quantity", "EnvQuantity",
                            "Center Env", "Total Env", "Env", "NRQuantity", "Nodal Code","Nodal Sort", "Route", "Route Sort", "Exam Time", "Exam Date", "BookletSerial","CourseName" };

                  var properties = new[] { "SerialNumber", "CatchNo", "CenterCode",
                               "CenterSort", "Quantity", "EnvQuantity",
                               "CenterEnv", "TotalEnv", "Env", "NRQuantity","NodalCode","NodalSort","Route","RouteSort", "ExamTime", "ExamDate","BookletSerial","CourseName" };

                  // Add filtered headers to the first row
                  for (int i = 0; i < headers.Length; i++)
                  {
                      worksheet.Cells[1, i + 1].Value = headers[i];
                  }

                  // Style headers
                  using (var range = worksheet.Cells[1, 1, 1, headers.Length])
                  {
                      range.Style.Font.Bold = true;
                      range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                      range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                      range.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                  }

                  // Add data rows dynamically
                  int row = 2;
                  foreach (Dictionary<string, object> rowItem in resultList)
                  {
                      string catchNo = rowItem.ContainsKey("CatchNo")
          ? rowItem["CatchNo"]?.ToString()
          : null;

                      // Reset OMR Serial if CatchNo changes and config enabled
                      if (resetOmrSerialOnCatchChange && prevCatchForSerial != null && catchNo != prevCatchForSerial)
                      {
                          currentStartNumber = startNumber;
                      }
                      bool isMssRow = rowItem.ContainsKey("isMss") && rowItem["isMss"] is bool bVal && bVal;
                      for (int col = 0; col < properties.Length; col++)
                      {
                          var propName = properties[col];

                          if (propName == "BookletSerial")
                          {
                              if (isMssRow)
                              {
                                  worksheet.Cells[row, col + 1].Value = "";
                                  continue;
                              }
                              if (assignBookletSerial)
                              {
                                  int envQuantity = 0;
                                  if (rowItem.ContainsKey("EnvQuantity") && rowItem["EnvQuantity"] != null)
                                      int.TryParse(rowItem["EnvQuantity"].ToString(), out envQuantity);

                                  string bookletSerialRange =
                                      $"{currentStartNumber}-{currentStartNumber + envQuantity - 1}";

                                  worksheet.Cells[row, col + 1].Value = bookletSerialRange;
                                  currentStartNumber += envQuantity;
                              }
                              else
                              {
                                  worksheet.Cells[row, col + 1].Value = "";
                              }
                          }
                          else
                          {
                              worksheet.Cells[row, col + 1].Value =
                                  rowItem.ContainsKey(propName)
                                  ? rowItem[propName]
                                  : null;
                          }
                      }
                      prevCatchForSerial = catchNo;
                      row++;
                  }


                  worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
                  worksheet.View.FreezePanes(2, 1);
                  // Save the file
                  var reportPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", ProjectId.ToString());
                  Directory.CreateDirectory(reportPath); // CreateDirectory is idempotent

                  var filePath = Path.Combine(reportPath, "EnvelopeBreaking.xlsx");
                  if (System.IO.File.Exists(filePath))
                  {
                      System.IO.File.Delete(filePath);
                  }

                  package.SaveAs(new FileInfo(filePath));
                  using var client = new HttpClient();
                  var response = await client.PostAsync(
       $"{_apiSettings.EnvelopeBreaking}?ProjectId={ProjectId}",
       new StringContent("") // required
   );
                  if (!response.IsSuccessStatusCode)
                  {
                      var error = await response.Content.ReadAsStringAsync();
                      Console.WriteLine($"API Failed: {response.StatusCode}, {error}");
                  }
                  else
                  {
                      var data = await response.Content.ReadAsStringAsync();
                      Console.WriteLine($"API Success: {data}");
                  }
                  return Ok(new { Result = resultList });
              }

          }*/

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
                _loggerService.LogEvent($"Deleted Envelope Breaking of Id {id}", "EnvelopeBreakages", LogHelper.GetTriggeredBy(User), envelopeBreakage.ProjectId);
                await _context.SaveChangesAsync();

                return NoContent();
            }
            catch (Exception ex)
            {
                _loggerService.LogError("Error deleting EnvelopeBreakage", ex.Message, nameof(EnvelopeBreakagesController));
                return StatusCode(500, "Internal Server Error");
            }

        }

        private bool EnvelopeBreakageExists(int id)
        {
            return _context.EnvelopeBreakages.Any(e => e.EnvelopeId == id);
        }
    }
}

