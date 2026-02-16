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

namespace Tools.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class EnvelopeBreakagesController : ControllerBase
    {
        private readonly ERPToolsDbContext _context;
        private readonly ILoggerService _loggerService;
        public EnvelopeBreakagesController(ERPToolsDbContext context, ILoggerService loggerService)
        {
            _context = context;
            _loggerService = loggerService;
        }

        // GET: api/EnvelopeBreakages
        [HttpGet]
        public async Task<ActionResult> GetEnvelopeBreakages(int ProjectId)
        {
            var NRData = await _context.NRDatas
                .Where(p => p.ProjectId == ProjectId)
                .ToListAsync();

            var Envelope = await _context.EnvelopeBreakages
                .Where(p => p.ProjectId == ProjectId)
                .ToListAsync();

            if (!NRData.Any() || !Envelope.Any())
                return NotFound("No data available for this project.");
            _loggerService.LogEvent($"No data available for this project", "EnvelopeBreakage", User.Identity?.Name != null ? int.Parse(User.Identity.Name) : 0, ProjectId);


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

            // 📁 Skip generation if file already exists
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

            // 🧾 Generate Excel report
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
                    package.SaveAs(new FileInfo(filePath));
                }
                _loggerService.LogEvent($"EnvelopeBreakage report of ProjectId {ProjectId} has been created", "EnvelopeBreakage", User.Identity?.Name != null ? int.Parse(User.Identity.Name) : 0, ProjectId);
                return Ok(Consolidated); // Return original data for UI (optional)
            }
            catch (Exception ex)
            {

                _loggerService.LogError("Error in generating report", ex.Message, nameof(EnvelopeBreakagesController));
                return StatusCode(500, "Internal server error");
            }

        }


        // GET: api/EnvelopeBreakages/5
        [HttpGet("{id}")]
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
                _loggerService.LogEvent($"Updated EnvelopeBreakage with ID {id}", "EnvelopeBreakage", User.Identity?.Name != null ? int.Parse(User.Identity.Name) : 0, envelopeBreakage.ProjectId);
                await _context.SaveChangesAsync();
            }
            catch (Exception ex)
            {
                if (!EnvelopeBreakageExists(id))
                {
                    _loggerService.LogEvent($"EnvelopeBreakage with ID {id} not found during updating", "EnvelopeBreakage", User.Identity?.Name != null ? int.Parse(User.Identity.Name) : 0, envelopeBreakage.ProjectId);
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
                    .Where(s => s.ProjectId == ProjectId)
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
                    // Determine the chunk size, adjust this based on your performance testing
                    var chunkSize = 1000; // Change this as necessary

                    // Loop through the list and remove in chunks
                    for (int i = 0; i < env.Count; i += chunkSize)
                    {
                        var chunk = env.Skip(i).Take(chunkSize).ToList();  // Get a chunk of the list
                        _context.EnvelopeBreakages.RemoveRange(chunk);  // Remove the chunk
                        await _context.SaveChangesAsync();  // Save changes after each chunk
                        _loggerService.LogEvent($"Deleted {chunk.Count} Envelope Breaking entries for ProjectID {ProjectId}", "EnvelopeBreakages",
                            User.Identity?.Name != null ? int.Parse(User.Identity.Name) : 0, ProjectId);
                    }

                    _loggerService.LogEvent($"Successfully deleted all Envelope Breaking entries for ProjectID {ProjectId}", "EnvelopeBreakages",
                        User.Identity?.Name != null ? int.Parse(User.Identity.Name) : 0, ProjectId);
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

                    remaining = quantity; // reset for outer
                    Dictionary<string, string> outerBreakdown = new();
                    int totalOuterCount = 0;

                    // Check if inner and outer are same
                    bool sameEnvelopes = innerSizes.SequenceEqual(outerSizes);

                    if (sameEnvelopes)
                    {
                        // Use same logic as inner
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
                    }
                    else
                    {
                        // Use original logic with rounding up
                        foreach (var size in outerSizes)
                        {
                            int count = (int)Math.Ceiling((double)remaining / size);
                            if (count > 0)
                            {
                                outerBreakdown[$"E{size}"] = count.ToString();
                                totalOuterCount += count;
                                remaining -= count * size;
                            }
                        }
                    }


                    var envelope = new EnvelopeBreakage
                    {
                        ProjectId = ProjectId,
                        NrDataId = row.Id,
                        InnerEnvelope = JsonSerializer.Serialize(innerBreakdown),
                        OuterEnvelope = JsonSerializer.Serialize(outerBreakdown),
                        TotalEnvelope = totalOuterCount
                    };

                    // ✅ Add to database

                    breakagesToAdd.Add(envelope);

                }

                if (breakagesToAdd.Any())
                {
                    _context.EnvelopeBreakages.AddRange(breakagesToAdd);
                    await _context.SaveChangesAsync();
                    _loggerService.LogEvent($"Created Envelope Breaking of ProjectID {ProjectId}", "EnvelopeBreakages", User.Identity?.Name != null ? int.Parse(User.Identity.Name) : 0, ProjectId);
                }


                using var client = new HttpClient();
/*                var response = await client.GetAsync($"http://192.168.10.208:81/API/api/EnvelopeBreakages/EnvelopeBreakage?ProjectId={ProjectId}");
*/                var response = await client.GetAsync($"https://localhost:7276/api/EnvelopeBreakages/EnvelopeBreakage?ProjectId={ProjectId}");

                if (!response.IsSuccessStatusCode)
                {
                    // Handle failure from GET call as needed
                    _loggerService.LogError("Failed to generate report", "", nameof(EnvelopeBreakagesController));
                    return StatusCode((int)response.StatusCode, "Failed to get envelope breakages after configuration.");
                }

                var jsonString = await response.Content.ReadAsStringAsync();
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




        [HttpGet("Replication")]
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
                .Where(p => p.ProjectId == ProjectId)
                .ToListAsync();

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
            // 🔹 Extract numeric part strictly
            var outerParts = outerEnvValue.Split(',', StringSplitOptions.RemoveEmptyEntries);

            // 🔹 Only parse if there’s a single non-empty value
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

            _loggerService.LogEvent($"Fieldnames  {string.Join(", ", fieldNames)}", "EnvelopeBreakages", User.Identity?.Name != null ? int.Parse(User.Identity.Name) : 0, ProjectId);
            var dupNames = await _context.Fields
                .Where(f => duplicatesFields.Contains(f.FieldId))
                .Select(f => f.Name)
                .ToListAsync();


            var startBox = ProjectConfig.BoxNumber;

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

            // 📁 Skip generation if file already exists
            if (System.IO.File.Exists(filePath))
            {
                System.IO.File.Delete(filePath);
            }
            // Define path to breakingreport.xlsx
            var breakingReportPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", ProjectId.ToString(), "BreakingReport.xlsx");

            // Check if file exists
            if (!System.IO.File.Exists(breakingReportPath))
            {
                return NotFound(new { message = "breakingreport.xlsx not found" });
            }

            var breakingReportData = new List<ExcelInputRow>();
            try
            {
                using (var package = new ExcelPackage(new FileInfo(breakingReportPath)))
                {
                    var worksheet = package.Workbook.Worksheets.FirstOrDefault();
                    if (worksheet == null)
                    {
                        return BadRequest(new { message = "No worksheet found in breakingreport.xlsx" });
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
                                ExamTime = worksheet.Cells[row, 14].Text.Trim(),
                                ExamDate = worksheet.Cells[row, 15].Text.Trim(),
                                Quantity = int.Parse(worksheet.Cells[row, 5].Text),
                                TotalEnv = int.Parse(worksheet.Cells[row, 8].Text),
                                NRQuantity = int.Parse(worksheet.Cells[row, 10].Text),
                                NodalCode = worksheet.Cells[row, 11].Text.Trim(), 
                                NodalSort = double.TryParse(worksheet.Cells[row, 12].Text.Trim(), out double sortVal)
                                ? sortVal
                                : 0.0,
                                RouteSort = Convert.ToInt32(worksheet.Cells[row, 13].Text.Trim()),
                                OmrSerial = worksheet.Cells[row,16].Text.Trim(),   
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
                        row.RouteSort,
                        Start = start,
                        End = end,
                        Serial = serial,
                        row.OmrSerial,
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
                    Property = enrichedList.First().GetType().GetProperty(name)
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
                        // If it’s numeric, fine — return it
                        if (double.TryParse(val.ToString(), out double nodalNum))
                            return nodalNum;

                        // ❌ Otherwise, throw to make the problem visible
                        throw new InvalidOperationException(
                            $"❌ NodalSort value is not numeric for record: {System.Text.Json.JsonSerializer.Serialize(x)} (actual value: '{val}')"
                        );
                    }

                    if (prop.Name.Equals("CenterSort", StringComparison.OrdinalIgnoreCase))
                    {
                        // If it’s numeric, fine — return it
                        if (int.TryParse(val.ToString(), out int centerNum))
                            return centerNum;

                        // ❌ Otherwise, throw to make the problem visible
                        throw new InvalidOperationException(
                            $"❌ CenterSort value is not numeric for record: {System.Text.Json.JsonSerializer.Serialize(x)} (actual value: '{val}')"
                        );
                    }
                    if (prop.Name.Equals("RouteSort", StringComparison.OrdinalIgnoreCase))
                    {
                        // If it’s numeric, fine — return it
                        if (int.TryParse(val.ToString(), out int routeNum))
                            return routeNum;

                        // ❌ Otherwise, throw to make the problem visible
                        throw new InvalidOperationException(
                            $"❌ RouteSort value is not numeric for record: {System.Text.Json.JsonSerializer.Serialize(x)} (actual value: '{val}')"
                        );
                    }
                    if (val is DateTime dt)
                        return dt;

                    // Try to parse as numeric if field name suggests it (ends with Sort or contains numeric indicators)
                  

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

            foreach (var item in sortedList)
            {
                var nrRow = nrData.FirstOrDefault(n => n.CatchNo == item.CatchNo);
                    int pages = nrRow?.Pages ?? 0;
                int totalPages = (item.Quantity) * pages;
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
                // ---- Rule 1: merge fields change → force new box
                bool mergeChanged = (prevMergeKey != null && mergeKey != prevMergeKey);

                try
                {
                    if (mergeChanged)
                    {
                        boxNo++; // start new box for new merge group
                        runningPages = 0;
                        _loggerService.LogEvent($"🔁 MergeKey changed → new box {boxNo} for {mergeKey}",
                            "EnvelopeBreakages",
                            User.Identity?.Name != null ? int.Parse(User.Identity.Name) : 0,
                            ProjectId);
                    }
                    bool overflow = (runningPages + totalPages > capacity);
                    if (overflow)
                    {
                        _loggerService.LogEvent($"Overflow {string.Join(", ", mergeKey)} Running {runningPages} Total Pages {totalPages}, Capacity {capacity} boxNo {boxNo}" +
                               $"", "EnvelopeBreakages", User.Identity?.Name != null ? int.Parse(User.Identity.Name) : 0, ProjectId);
                        Console.WriteLine(envelopeSize);
                        if (envelopeSize < 50 && envelopeSize>0)
                        {
                            _loggerService.LogEvent($"MergeKey overflow {string.Join(", ", mergeKey)} Running {runningPages} Total Pages {totalPages}, Capacity {capacity} boxNo {boxNo}" +
                                $"", "EnvelopeBreakages", User.Identity?.Name != null ? int.Parse(User.Identity.Name) : 0, ProjectId);
                            int overflowAmount = (runningPages + totalPages) - capacity;
                            int pagesPerUnit = (nrRow?.Pages ?? 0);

                            // 🔹 Calculate total pages including current box content
                            int totalPagesAll = runningPages + totalPages;

                            // 🔹 How many boxes are needed to hold all pages
                            int boxesNeeded = (int)Math.Ceiling((double)totalPagesAll / capacity);

                            // 🔹 Divide quantity into these boxes
                            int totalQty = item.Quantity;
                            int baseQty = (int)Math.Floor((double)totalQty / boxesNeeded);

                            // Round each box's quantity to nearest multiple of 20
                            baseQty = (int)Math.Round(baseQty / (double)envelopeSize) * envelopeSize;
                            if (baseQty <= 0) baseQty = envelopeSize; // safety floor

                            // 🔹 Now distribute across boxes
                            int distributedQty = 0;
                            int currentStart = 1;

                            // if we already have boxes before, continue from last box's end + 1
                            // if we already have boxes for this CatchNo, continue from last box's end + 1
                            var lastBoxForCatch = finalWithBoxes
                                .Where(b => (string)b.GetType().GetProperty("CatchNo").GetValue(b) == item.CatchNo)
                                .OrderBy(b => (int)b.GetType().GetProperty("End").GetValue(b))
                                .LastOrDefault();

                            if (lastBoxForCatch != null)
                            {
                                currentStart = (int)lastBoxForCatch.GetType().GetProperty("End").GetValue(lastBoxForCatch) + 1;
                            }
                            else
                            {
                                currentStart = 1; // start fresh if no previous box exists for this CatchNo
                            }


                            for (int b = 1; b <= boxesNeeded; b++)
                            {
                                if (b > 1) boxNo++;

                                int boxQty;
                                if (b == boxesNeeded)
                                {
                                    // last box gets remainder
                                    boxQty = totalQty - distributedQty;
                                }
                                else
                                {
                                    boxQty = baseQty;
                                    distributedQty += boxQty;
                                }

                                // ensure no negative/remainder issues
                                if (boxQty <= 0) continue;

                                int boxPages = boxQty * pagesPerUnit;
                                int envelopesInBox = boxQty / envelopeSize;
                                int start = currentStart;
                                int end = start + envelopesInBox - 1;
                                currentStart = end + 1;
                                string serial = $"{start} to {end}";
                                finalWithBoxes.Add(new
                                {
                                    item.CatchNo,
                                    item.CenterCode,
                                    item.CenterSort,
                                    item.ExamTime,
                                    item.ExamDate,
                                    Quantity = boxQty,
                                    item.NodalCode,
                                    item.NodalSort,
                                    item.RouteSort,
                                    item.TotalEnv,
                                    Start = start,
                                    End = end,
                                    Serial = serial,
                                    TotalPages = boxPages,
                                    BoxNo = boxNo,
                                    item.OmrSerial,
                                });
                            }

                            // ✅ reset running values
                            runningPages = 0;
                            prevMergeKey = mergeKey;
                            continue;
                        }
                        boxNo++;
                        runningPages = 0;
                    }

                  
                        // normal case: just start new box
                    runningPages += totalPages;

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
                        item.RouteSort,
                        item.TotalEnv,
                        item.Start,
                        item.End,
                        item.Serial,
                        TotalPages = totalPages,
                        BoxNo = boxNo,
                        item.OmrSerial,
                    });

                    prevMergeKey = mergeKey;
                }

                catch (Exception ex)
                {
                    _loggerService.LogError("Error storing report EnvelopeBreakage", ex.Message, nameof(EnvelopeBreakagesController));
                    return StatusCode(500, "Internal Server Error");
                }
            }

            // 🔹 Maintain ordering
            finalWithBoxes = finalWithBoxes
                .OrderBy(x => (int)x.BoxNo)
                .ToList();


            // Step 5: Export to Excel
            try
            {
                using (var package = new ExcelPackage())
                {
                    var worksheet = package.Workbook.Worksheets.Add("ReplicationResult");

                    // Headers
                    worksheet.Cells[1, 1].Value = "SerialNumber";
                    worksheet.Cells[1, 2].Value = "CatchNo";
                    worksheet.Cells[1, 3].Value = "CenterCode";
                    worksheet.Cells[1, 4].Value = "CenterSort";
                    worksheet.Cells[1, 5].Value = "ExamTime";
                    worksheet.Cells[1, 6].Value = "ExamDate";
                    worksheet.Cells[1, 7].Value = "Quantity";
                    worksheet.Cells[1, 8].Value = "NodalCode"; 
                    worksheet.Cells[1, 9].Value = "NodalSort";
                    worksheet.Cells[1, 10].Value = "RouteSort";
                    worksheet.Cells[1, 11].Value = "TotalEnv";
                    worksheet.Cells[1, 12].Value = "Start";
                    worksheet.Cells[1, 13].Value = "End";
                    worksheet.Cells[1, 14].Value = "Serial";
                    worksheet.Cells[1, 15].Value = "TotalPages";
                    worksheet.Cells[1, 16].Value = "BoxNo";
                    worksheet.Cells[1, 17].Value = "OmrSerial";
                    int row = 2;
                    int serial = 1;
                    foreach (var item in finalWithBoxes)
                    {
                        worksheet.Cells[row, 1].Value = serial++;
                        worksheet.Cells[row, 2].Value = item.CatchNo;
                        worksheet.Cells[row, 3].Value = item.CenterCode; 
                        worksheet.Cells[row, 4].Value = item.CenterSort;
                        worksheet.Cells[row, 5].Value = item.ExamTime;
                        worksheet.Cells[row, 6].Value = item.ExamDate;
                        worksheet.Cells[row, 7].Value = item.Quantity;
                        worksheet.Cells[row, 8].Value = item.NodalCode;
                        worksheet.Cells[row, 9].Value = item.NodalSort;
                        worksheet.Cells[row, 10].Value = item.RouteSort;
                        worksheet.Cells[row, 11].Value = item.TotalEnv;
                        worksheet.Cells[row, 12].Value = item.Start;
                        worksheet.Cells[row, 13].Value = item.End;
                        worksheet.Cells[row, 14].Value = item.Serial;
                        worksheet.Cells[row, 15].Value = item.TotalPages;
                        worksheet.Cells[row, 16].Value = item.BoxNo;
                        worksheet.Cells[row, 17].Value = item.OmrSerial;
                        row++;
                    }

                    FileInfo fi = new FileInfo(filePath);
                    package.SaveAs(fi);
                }

                return Ok(new { message = "File successfully created", filePath });
            }
            catch (Exception ex)
            {
                _loggerService.LogError("Error creating report EnvelopeBreakage", ex.Message, nameof(EnvelopeBreakagesController));
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
            public int RouteSort { get; set; }
            public string OmrSerial { get; set; }
        }


        [HttpGet("EnvelopeBreakage")]
        public async Task<IActionResult> BreakageConfiguration(int ProjectId)
        {
            // Fetch all data sequentially
            var envCaps = await _context.EnvelopesTypes
                .Select(e => new { e.EnvelopeName, e.Capacity })
                .ToListAsync();

            var nrData = await _context.NRDatas
                .Where(p => p.ProjectId == ProjectId)
                .OrderBy(p=>p.CatchNo)
                .ThenBy(p => p.RouteSort)
                .ThenBy(p => p.NodalSort)
                .ThenBy(p => p.CenterSort)
                .ToListAsync();
         

            var envBreaking = await _context.EnvelopeBreakages
                .Where(p => p.ProjectId == ProjectId)
                .ToListAsync();

            var extras = await _context.ExtrasEnvelope
                .Where(p => p.ProjectId == ProjectId)
                .ToListAsync();

            var projectconfig = await _context.ProjectConfigs
                .Where(p => p.ProjectId == ProjectId)
                .FirstOrDefaultAsync();

            var outerEnvJson = projectconfig.Envelope;

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
               e => (e.TotalEnvelope, e.OuterEnvelope) // 👈 this is a tuple
            );


            var resultList = new List<object>();
            int globalSerialNumber = 1; // Global serial number that only resets on CatchNo change
            string prevNodalCode = null;
            int prevNodalSort = 0;
            string prevCatchNo = null;
            string prevMergeField = null;
            string prevExtraMergeField = null;
            int centerEnvCounter = 0;
            int extraCenterEnvCounter = 0;
            var nodalExtrasAddedForNodalCatch = new HashSet<(string NodalCode, string CatchNo)>();
            var catchExtrasAdded = new HashSet<(int ExtraId, string CatchNo)>();

            // Helper method to add extra envelopes - removed serialnumber++ from here
            void AddExtraWithEnv(ExtraEnvelopes extra, string examDate, string examTime, int NrQuantity, string NodalCode, string CenterCode, int CenterSort, int NodalSort)
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
                        SerialNumber = globalSerialNumber++, // Use global serial number
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
                        ExamTime = examTime,
                        TotalEnv = totalEnv,
                        Env = $"{j}/{totalEnv}",
                        NRQuantity = NrQuantity,
                        NodalCode = NodalCode,
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
                              1 => 1000,
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
                    // ➕ Add final extras for previous CatchNo before resetting serial
                    if (!nodalExtrasAddedForNodalCatch.Contains((prevNrData.NodalCode, prevCatchNo)))
                    {
                        var extrasToAdd = extras.Where(e => e.ExtraId == 1 && e.CatchNo == prevCatchNo).ToList();
                        // Get previous record for metadata
                        foreach (var extra in extrasToAdd)
                        {
                            AddExtraWithEnv(extra, prevNrData.ExamDate, prevNrData.ExamTime,
                                          prevNrData.NRQuantity, prevNrData.NodalCode, prevNrData.CenterCode, prevNrData.CenterSort, prevNrData.NodalSort);
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
                                AddExtraWithEnv(extra, prevNrData.ExamDate, prevNrData.ExamTime,
                                               prevNrData.NRQuantity, prevNrData.NodalCode, prevNrData.CenterCode, prevNrData.CenterSort, prevNrData.NodalSort);
                            }
                            catchExtrasAdded.Add((extraId, prevCatchNo));
                        }
                    }

                    // NOW reset serial number after extras are added
                    globalSerialNumber = 1;
                }

                // ➕ Nodal Extra when NodalCode changes (but not CatchNo)
                if (!catchNoChanged && prevNodalCode != null && current.NodalCode != prevNodalCode)
                {
                    Console.WriteLine($" Processing: {nrData[i].CatchNo}, CenterSort={nrData[i].CenterSort}, NodalSort={nrData[i].NodalSort}");

                    if (!nodalExtrasAddedForNodalCatch.Contains((prevNodalCode, current.CatchNo)))
                    {
                        var extrasToAdd = extras.Where(e => e.ExtraId == 1 && e.CatchNo == current.CatchNo).ToList();
                        foreach (var extra in extrasToAdd)
                        {
                            AddExtraWithEnv(extra, current.ExamDate, current.ExamTime,
                                           current.NRQuantity, prevNodalCode, current.CenterCode, current.CenterSort, prevNodalSort);
                        }
                        nodalExtrasAddedForNodalCatch.Add((prevNodalCode, current.CatchNo));
                    }
                }

                // ➕ Add current NRData row with TotalEnv replication
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
                                SerialNumber = globalSerialNumber++,
                                current.CatchNo,
                                current.CenterCode,
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
                                current.RouteSort,
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
                prevCatchNo = current.CatchNo;
            }

            // 🔁 Final extras for the last CatchNo
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
                            AddExtraWithEnv(extra, lastNrData.ExamDate, lastNrData.ExamTime,
                                       lastNrData.NRQuantity, lastNrData.NodalCode, lastNrData.CenterCode, lastNrData.CenterSort, lastNrData.NodalSort);
                        }
                    }

                    foreach (var extraId in new[] { 2, 3 })
                    {
                        if (!catchExtrasAdded.Contains((extraId, prevCatchNo)))
                        {
                            var extrasToAdd = extras.Where(e => e.ExtraId == extraId && e.CatchNo == prevCatchNo).ToList();
                            foreach (var extra in extrasToAdd)
                            {
                                AddExtraWithEnv(extra, lastNrData.ExamDate, lastNrData.ExamTime,
                                             lastNrData.NRQuantity, lastNrData.NodalCode, lastNrData.CenterCode,lastNrData.CenterSort,lastNrData.NodalSort);
                            }
                        }
                    }
                }
            }
            int currentStartNumber = startNumber;
            bool assignBookletSerial = currentStartNumber > 0;
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
                    ordered = resultList.OrderBy(keySelector);
                else
                    ordered = ordered.ThenBy(keySelector);
            }

            resultList = ordered?.ToList() ?? resultList;

            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("BreakingResult");

                // Add headers
                var headers = new[] { "Serial Number", "Catch No", "Center Code",
                        "Center Sort", "Quantity", "EnvQuantity",
                          "Center Env", "Total Env", "Env", "NRQuantity", "Nodal Code","Nodal Sort", "Route Sort", "Exam Time", "Exam Date", "BookletSerial" };

                var properties = new[] { "SerialNumber", "CatchNo", "CenterCode",
                             "CenterSort", "Quantity", "EnvQuantity",
                             "CenterEnv", "TotalEnv", "Env", "NRQuantity","NodalCode","NodalSort","RouteSort", "ExamTime", "ExamDate","BookletSerial" };

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
                foreach (var item in resultList)
                {
                    dynamic rowItem = item;
                    for (int col = 0; col < properties.Length; col++)
                    {
                        var propName = properties[col];

                        if (propName == "BookletSerial")
                        {
                            if (assignBookletSerial)
                            {
                                int envQuantity = rowItem.EnvQuantity;
                                string bookletSerialRange = $"{currentStartNumber}-{currentStartNumber + envQuantity - 1}";
                                worksheet.Cells[row, col + 1].Value = bookletSerialRange;
                                currentStartNumber += envQuantity;
                            }
                            else
                            {
                                worksheet.Cells[row, col + 1].Value = ""; // Leave it blank
                            }
                        }
                        else
                        {
                            var value = rowItem.GetType().GetProperty(propName)?.GetValue(rowItem);
                            worksheet.Cells[row, col + 1].Value = value;
                        }
                    }
                    row++;
                }

                worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();

                // Save the file
                var reportPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", ProjectId.ToString());
                Directory.CreateDirectory(reportPath); // CreateDirectory is idempotent

                var filePath = Path.Combine(reportPath, "BreakingReport.xlsx");
                if (System.IO.File.Exists(filePath))
                {
                    System.IO.File.Delete(filePath);
                }

                package.SaveAs(new FileInfo(filePath));

                return Ok(new { Result = resultList });
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
                _loggerService.LogEvent($"Deleted Envelope Breaking of Id {id}", "EnvelopeBreakages", User.Identity?.Name != null ? int.Parse(User.Identity.Name) : 0, envelopeBreakage.ProjectId);
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
