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

namespace Tools.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class EnvelopeBreakagesController : ControllerBase
    {
        private readonly ERPToolsDbContext _context;

        public EnvelopeBreakagesController(ERPToolsDbContext context)
        {
            _context = context;
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
                    catch { /* Ignore */ }
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
                    catch { }
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
                    catch { }
                }

                parsedRows.Add(parsedRow);
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

            return Ok(Consolidated); // Return original data for UI (optional)
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
                await _context.SaveChangesAsync();
            }
            catch (DbUpdateConcurrencyException)
            {
                if (!EnvelopeBreakageExists(id))
                {
                    return NotFound();
                }
                else
                {
                    throw;
                }
            }

            return NoContent();
        }

        // POST: api/EnvelopeBreakages
        // To protect from overposting attacks, see https://go.microsoft.com/fwlink/?linkid=2123754
        [HttpPost("EnvelopeConfiguration")]
        public async Task<IActionResult> EnvelopeConfiguration(int ProjectId)
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
                _context.EnvelopeBreakages.RemoveRange(env);  // Use RemoveRange without Async
                await _context.SaveChangesAsync();  // Save changes asynchronously
            }
            else
            {
                // Handle the case where no records were found
                Console.WriteLine("No breakages found for the specified project.");
            }

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

                var envelope = new EnvelopeBreakage
                {
                    ProjectId = ProjectId,
                    NrDataId = row.Id,
                    InnerEnvelope = JsonSerializer.Serialize(innerBreakdown),
                    OuterEnvelope = JsonSerializer.Serialize(outerBreakdown),
                    TotalEnvelope = totalOuterCount
                };

                // ✅ Add to database
                _context.EnvelopeBreakages.Add(envelope);
            }

            await _context.SaveChangesAsync();

            using var client = new HttpClient();
            var response = await client.GetAsync($"http://192.168.10.208:81/API/api/EnvelopeBreakages/EnvelopeBreakage?ProjectId={ProjectId}");

            if (!response.IsSuccessStatusCode)
            {
                // Handle failure from GET call as needed
                return StatusCode((int)response.StatusCode, "Failed to get envelope breakages after configuration.");
            }

            var jsonString = await response.Content.ReadAsStringAsync();


            return Ok("Envelope breakdown successfully saved to EnvelopeBreakage table.");
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

            var capacity = await _context.ProjectConfigs
                .Where (p => p.ProjectId == ProjectId)
                .Select (p => p.BoxCapacity)
                .FirstOrDefaultAsync();

 
            var boxIds = await _context.ProjectConfigs
                .Where (p => p.ProjectId == ProjectId)
                .Select (p => p.BoxBreakingCriteria) .FirstOrDefaultAsync();

            var fields = await _context.Fields
                .Where(f=>boxIds.Contains(f.FieldId))
                .ToListAsync();

            // Step 2: Fetch the corresponding field names from the Fields table
            var fieldNames = await _context.Fields
                .Where(f => boxIds.Contains(f.FieldId))  // Assuming envelopeIds contains the IDs of the fields to sort by
                .Select(f => f.Name)  // Get the field names
                .ToListAsync();


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
                            SerialNumber = int.Parse(worksheet.Cells[row, 1].Text),
                            CatchNo = worksheet.Cells[row, 2].Text.Trim(),
                            CenterCode = worksheet.Cells[row, 3].Text.Trim(),
                            ExamTime = worksheet.Cells[row, 4].Text.Trim(),
                            ExamDate = worksheet.Cells[row, 5].Text.Trim(),
                            Quantity = int.Parse(worksheet.Cells[row, 6].Text),
                            TotalEnv = int.Parse(worksheet.Cells[row, 9].Text),
                            NRQuantity = int.Parse(worksheet.Cells[row, 11].Text),
                            NodalCode = worksheet.Cells[row,12].Text.Trim(),
                        };

                        breakingReportData.Add(inputRow);
                    }
                    catch(Exception ex)
                    {
                        Console.WriteLine($"Error parsing row {row}: {ex.Message}");
                    }
                }
            }

            // Step 1: Remove duplicates (CatchNo + CenterCode), preserving first occurrence
            var uniqueRows = breakingReportData
            .GroupBy(x => $"{x.CatchNo}_{x.CenterCode}")
             .Select(g => g.First())
            .ToList();
           
           
            // Step 2: Calculate Start, End, Serial (before sorting)
            var enrichedList = new List<dynamic>();
            string previousCatchNo = null;
            int previousEnd = 0;

            foreach (var row in uniqueRows)
            {
                int start = row.CatchNo != previousCatchNo ? 1 : previousEnd + 1;
                int end = start + row.TotalEnv - 1;
                string serial = $"{start} to {end}";

                enrichedList.Add(new
                {
                    row.SerialNumber,
                    row.CatchNo,
                    row.CenterCode,
                    row.ExamTime,
                    row.ExamDate,
                    row.Quantity,
                    row.TotalEnv,
                    row.NRQuantity,
                    row.NodalCode,
                    Start = start,
                    End = end,
                    Serial = serial
                });

                previousCatchNo = row.CatchNo;
                previousEnd = end;
            }

            if (!enrichedList.Any())
                return Ok(new { message = "No data to process" });
            Console.WriteLine(fieldNames);
            // Step 1: Cache property info once
            var properties = fieldNames
                .Select(name => new
                {
                    Name = name,
                    Property = enrichedList.First().GetType().GetProperty(name)
                })
                .Where(x => x.Property != null)
                .ToList();
            foreach (var prop in properties)
            {
                Console.WriteLine($"Name: {prop.Name}, Property: {prop.Property.Name}");
            }

            // Step 2: Apply ordering using cached properties
            IOrderedEnumerable<dynamic> ordered = null;

            for (int i = 0; i < properties.Count; i++)
            {
                var prop = properties[i].Property;

                if (i == 0)
                    ordered = enrichedList.OrderBy(x => prop.GetValue(x));
                else
                    ordered = ordered.ThenBy(x => prop.GetValue(x));
            }

            // Step 3: Update the list if sorting happened
            if (ordered != null)
                enrichedList = ordered.ToList();

            var sortedList = ordered?.ToList() ?? enrichedList;
            // Step 6: Add TotalPages and BoxNo
            int boxNo = 1001;
            int runningPages = 0;
            string prevMergeKey = null;
            
            var finalWithBoxes = new List<dynamic>();

            foreach (var item in sortedList)
            {
                var nrRow = nrData.FirstOrDefault(n => n.CenterCode == item.CenterCode && n.CatchNo == item.CatchNo);
                int pages = nrRow?.Pages ?? 1;
                int totalPages = (item.Quantity ?? 0) * pages;

                // Build merge key
                string mergeKey = "";
                if (boxIds.Any())
                {
                    mergeKey = string.Join("_", boxIds.Select(fieldId =>
                    {
                        // Fetch the field name from Fields table (assumes fieldId is a valid FieldId)
                        var fieldName = fields.FirstOrDefault(f => f.FieldId == fieldId)?.Name;

                        // If the field exists in fields, get the corresponding value from nrRow
                        if (fieldName != null)
                        {
                            var prop = nrRow?.GetType().GetProperty(fieldName);
                            return prop?.GetValue(nrRow)?.ToString() ?? "";
                        }

                        return "";
                    }));
                }

                // ---- Rule 1: merge fields change → force new box
                bool mergeChanged = (prevMergeKey != null && mergeKey != prevMergeKey);

                // ---- Rule 2: page overflow
                bool overflow = (runningPages + totalPages > capacity);

                if (mergeChanged || overflow)
                {
                    // case: overflow detected
                    if (overflow)
                    {
                        int leftover = (runningPages + totalPages) - capacity;
                        int filled = totalPages - leftover;

                        // if leftover < 50% threshold, split into two balanced boxes
                        if (leftover > 0 && leftover < capacity/2)
                        {
                            int combined = capacity + leftover;
                            int half = combined / 2;

                            // Adjust: previous box gets half, new box gets half
                            // 💡 This means we must "reassign" some pages from last box items into new box
                            // Simplify: we treat this item as split across two boxes
                            // first portion
                            finalWithBoxes.Add(new
                            {
                                item.SerialNumber,
                                item.CatchNo,
                                item.CenterCode,
                                item.ExamTime,
                                item.ExamDate,
                                item.Quantity,
                                item.NodalCode,
                                item.TotalEnv,
                                item.Start,
                                item.End,
                                item.Serial,
                                TotalPages = half,
                                BoxNo = boxNo
                            });

                            // second portion
                            boxNo++;
                            finalWithBoxes.Add(new
                            {
                                item.SerialNumber,
                                item.CatchNo,
                                item.CenterCode,
                                item.ExamTime,
                                item.ExamDate,
                                item.Quantity,
                                item.NodalCode,
                                item.TotalEnv,
                                item.Start,
                                item.End,
                                item.Serial,
                                TotalPages = combined - half,
                                BoxNo = boxNo
                            });

                            runningPages = combined - half; // carry over for next items
                            prevMergeKey = mergeKey;
                            continue; // skip normal add
                        }
                    }

                    // normal case: just start new box
                    boxNo++;
                    runningPages = 0;
                }

                runningPages += totalPages;

                finalWithBoxes.Add(new
                {
                    item.SerialNumber,
                    item.CatchNo,
                    item.CenterCode,
                    item.ExamTime,
                    item.ExamDate,
                    item.Quantity,
                    item.NodalCode,
                    item.TotalEnv,
                    item.Start,
                    item.End,
                    item.Serial,
                    TotalPages = totalPages,
                    BoxNo = boxNo
                });

                prevMergeKey = mergeKey;
            }


            // Step 5: Export to Excel
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("ReplicationResult");

                // Headers
                worksheet.Cells[1, 1].Value = "SerialNumber";
                worksheet.Cells[1, 2].Value = "CatchNo";
                worksheet.Cells[1, 3].Value = "CenterCode";
                worksheet.Cells[1, 4].Value = "ExamTime";
                worksheet.Cells[1, 5].Value = "ExamDate";
                worksheet.Cells[1, 6].Value = "Quantity";
                worksheet.Cells[1, 7].Value = "NodalCode";
                worksheet.Cells[1, 8].Value = "TotalEnv";
                worksheet.Cells[1, 10].Value = "Start";
                worksheet.Cells[1, 11].Value = "End";
                worksheet.Cells[1, 12].Value = "Serial";
                worksheet.Cells[1, 13].Value = "TotalPages";
                worksheet.Cells[1, 14].Value = "BoxNo";

                int row = 2;
                foreach (var item in finalWithBoxes)
                {
                    worksheet.Cells[row, 1].Value = item.SerialNumber;
                    worksheet.Cells[row, 2].Value = item.CatchNo;
                    worksheet.Cells[row, 3].Value = item.CenterCode;
                    worksheet.Cells[row, 4].Value = item.ExamTime;
                    worksheet.Cells[row, 5].Value = item.ExamDate;
                    worksheet.Cells[row, 6].Value = item.Quantity;
                    worksheet.Cells[row, 7].Value = item.NodalCode;
                    worksheet.Cells[row, 8].Value = item.TotalEnv;
                    worksheet.Cells[row, 10].Value = item.Start;
                    worksheet.Cells[row, 11].Value = item.End;
                    worksheet.Cells[row, 12].Value = item.Serial;
                    worksheet.Cells[row, 13].Value = item.TotalPages;
                    worksheet.Cells[row, 14].Value = item.BoxNo;
                    row++;
                }

                FileInfo fi = new FileInfo(filePath);
                package.SaveAs(fi);
            }

            return Ok(new { message = "File successfully created", filePath });


        }

        public class ExcelInputRow
        {
            public int SerialNumber { get; set; }
            public string CatchNo { get; set; }
            public string CenterCode { get; set; }
            public string ExamTime { get; set; }
            public string ExamDate { get; set; }
            public int Quantity { get; set; }
            public int TotalEnv { get; set; }
            public int NRQuantity { get; set; }
            public string NodalCode { get; set; }
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
                .ToListAsync();

            var envBreaking = await _context.EnvelopeBreakages
                .Where(p => p.ProjectId == ProjectId)
                .ToListAsync();

            var extras = await _context.ExtrasEnvelope
                .Where(p => p.ProjectId == ProjectId)
                .ToListAsync();

            var outerEnvJson = await _context.ProjectConfigs
                .Where(p => p.ProjectId == ProjectId)
                .Select(p => p.Envelope)
                .FirstOrDefaultAsync();

            var extrasconfig = await _context.ExtraConfigurations
                .Where(p => p.ProjectId == ProjectId)
                .ToListAsync();

            // Build dictionaries for fast lookup
            var envelopeCapacities = envCaps.ToDictionary(x => x.EnvelopeName, x => x.Capacity);
            var envDict = envBreaking.ToDictionary(e => e.NrDataId, e => e.TotalEnvelope);

            var resultList = new List<object>();
            int globalSerialNumber = 1; // Global serial number that only resets on CatchNo change
            string prevNodalCode = null;
            string prevCatchNo = null;
            string prevMergeField = null;
            string prevExtraMergeField = null;
            int centerEnvCounter = 0;
            int extraCenterEnvCounter = 0;
            var nodalExtrasAddedForCatchNo = new HashSet<string>();
            var catchExtrasAdded = new HashSet<(int ExtraId, string CatchNo)>();

            // Extract outer envelope config
            int nrOuterCapacity = 100; // default
            if (!string.IsNullOrEmpty(outerEnvJson))
            {
                var outerEnvDict = JsonSerializer.Deserialize<Dictionary<string, string>>(outerEnvJson);
                if (outerEnvDict != null && outerEnvDict.TryGetValue("Outer", out string outerType)
                    && envelopeCapacities.TryGetValue(outerType, out int capacity))
                {
                    nrOuterCapacity = capacity;
                }
            }

            // Helper method to add extra envelopes - removed serialnumber++ from here
            void AddExtraWithEnv(ExtraEnvelopes extra, string examDate, string examTime, string course, string subject, int NrQuantity, string NodalCode)
            {
                var extraConfig = extrasconfig.FirstOrDefault(e => e.ExtraType == extra.ExtraId);
                int envCapacity = 100; // default fallback

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
                        CourseName = course,
                        SubjectName = subject,
                        TotalEnv = totalEnv,
                        Env = $"{j}/{totalEnv}",
                        NRQuantity = NrQuantity,
                        NodalCode = NodalCode,
                    });
                }
            }

            for (int i = 0; i < nrData.Count; i++)
            {
                var current = nrData[i];

                // Check if CatchNo changed BEFORE processing extras
                bool catchNoChanged = prevCatchNo != null && current.CatchNo != prevCatchNo;

                if (catchNoChanged)
                {
                    // ➕ Add final extras for previous CatchNo before resetting serial
                    if (!nodalExtrasAddedForCatchNo.Contains(prevCatchNo))
                    {
                        var extrasToAdd = extras.Where(e => e.ExtraId == 1 && e.CatchNo == prevCatchNo).ToList();
                        var prevNrData = nrData[i - 1]; // Get previous record for metadata
                        foreach (var extra in extrasToAdd)
                        {
                            AddExtraWithEnv(extra, prevNrData.ExamDate, prevNrData.ExamTime,
                                          prevNrData.SubjectName, prevNrData.CourseName, prevNrData.NRQuantity, prevNrData.NodalCode);
                        }
                        nodalExtrasAddedForCatchNo.Add(prevCatchNo);
                    }

                    foreach (var extraId in new[] { 2, 3 })
                    {
                        if (!catchExtrasAdded.Contains((extraId, prevCatchNo)))
                        {
                            var extrasToAdd = extras.Where(e => e.ExtraId == extraId && e.CatchNo == prevCatchNo).ToList();
                            var prevNrData = nrData[i - 1];
                            foreach (var extra in extrasToAdd)
                            {
                                AddExtraWithEnv(extra, prevNrData.ExamDate, prevNrData.ExamTime,
                                              prevNrData.SubjectName, prevNrData.CourseName, prevNrData.NRQuantity,prevNrData.NodalCode);
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
                    if (!nodalExtrasAddedForCatchNo.Contains(current.CatchNo))
                    {
                        var extrasToAdd = extras.Where(e => e.ExtraId == 1 && e.CatchNo == current.CatchNo).ToList();
                        foreach (var extra in extrasToAdd)
                        {
                            AddExtraWithEnv(extra, current.ExamDate, current.ExamTime,
                                          current.SubjectName, current.CourseName, current.NRQuantity,current.NodalCode);
                        }
                        nodalExtrasAddedForCatchNo.Add(current.CatchNo);
                    }
                }

                // ➕ Add current NRData row with TotalEnv replication
                int totalEnv = envDict.TryGetValue(current.Id, out int envCount) && envCount > 0 ? envCount : 1;

                // Calculate CenterEnv
                string currentMergeField = $"{current.CatchNo}-{current.CenterCode}";
                if (currentMergeField != prevMergeField)
                {
                    centerEnvCounter = 0;
                    prevMergeField = currentMergeField;
                }

                for (int j = 1; j <= totalEnv; j++)
                {
                    centerEnvCounter++;
                    int envQuantity;

                    if (current.Quantity > nrOuterCapacity && j == 1)
                    {
                        envQuantity = current.Quantity - (totalEnv - 1) * nrOuterCapacity;
                    }
                    else if (current.Quantity <= nrOuterCapacity)
                    {
                        envQuantity = current.Quantity;
                    }
                    else
                    {
                        envQuantity = nrOuterCapacity;
                    }

                    resultList.Add(new
                    {
                        SerialNumber = globalSerialNumber++,
                        current.CourseName,
                        current.SubjectName,
                        current.CatchNo,
                        current.CenterCode,
                        current.ExamTime,
                        current.ExamDate,
                        current.Quantity,
                        EnvQuantity = envQuantity,
                        current.NodalCode,
                        CenterEnv = centerEnvCounter,
                        TotalEnv = totalEnv,
                        Env = $"{centerEnvCounter}/{totalEnv}",
                        current.NRQuantity,
                    });
                }

                prevNodalCode = current.NodalCode;
                prevCatchNo = current.CatchNo;
            }

            // 🔁 Final extras for the last CatchNo
            if (prevCatchNo != null)
            {
                var lastNrData = nrData.LastOrDefault();

                if (lastNrData != null)
                {
                    if (!nodalExtrasAddedForCatchNo.Contains(prevCatchNo))
                    {
                        var extrasToAdd = extras.Where(e => e.ExtraId == 1 && e.CatchNo == prevCatchNo).ToList();
                        foreach (var extra in extrasToAdd)
                        {
                            AddExtraWithEnv(extra, lastNrData.ExamDate, lastNrData.ExamTime,
                                          lastNrData.SubjectName, lastNrData.CourseName, lastNrData.NRQuantity,lastNrData.NodalCode);
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
                                              lastNrData.SubjectName, lastNrData.CourseName, lastNrData.NRQuantity,lastNrData.NodalCode);
                            }
                        }
                    }
                }
            }

            // Generate Excel Report
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("BreakingResult");

                // Add headers
                var headers = new[] { "Serial Number", "Course Name", "Subject Name", "Catch No", "Center Code",
                          "Exam Time", "Exam Date", "Quantity", "EnvQuantity", "Nodal Code",
                          "Center Env", "Total Env", "Env", "NRQuantity" };

                var properties = new[] { "SerialNumber", "CourseName", "SubjectName", "CatchNo", "CenterCode",
                            "ExamTime", "ExamDate", "Quantity", "EnvQuantity", "NodalCode",
                             "CenterEnv", "TotalEnv", "Env", "NRQuantity" };

                // Create a list to track which columns should be included (those that contain data)
                var columnsToInclude = new List<int>();

                // Check for non-empty columns in any row (resultList) and track their indices
                foreach (var item in resultList)
                {
                    for (int i = 0; i < properties.Length; i++)
                    {
                        var value = item.GetType().GetProperty(properties[i])?.GetValue(item);
                        if (value != null && !string.IsNullOrEmpty(value.ToString()))
                        {
                            if (!columnsToInclude.Contains(i))
                            {
                                columnsToInclude.Add(i);  // Add column index if it contains data
                            }
                        }
                    }
                }

                // Filter headers and properties based on the columns that have data
                var filteredHeaders = columnsToInclude.Select(index => headers[index]).ToArray();
                var filteredProperties = columnsToInclude.Select(index => properties[index]).ToArray();

                // Add filtered headers to the first row
                for (int i = 0; i < filteredHeaders.Length; i++)
                {
                    worksheet.Cells[1, i + 1].Value = filteredHeaders[i];
                }

                // Style headers
                using (var range = worksheet.Cells[1, 1, 1, filteredHeaders.Length])
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
                    for (int col = 0; col < filteredProperties.Length; col++)
                    {
                        var value = item.GetType().GetProperty(filteredProperties[col])?.GetValue(item);
                        worksheet.Cells[row, col + 1].Value = value;
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

        /*  [HttpGet("EnvelopeBreakage")]
          public async Task<IActionResult> BreakageConfiguration(int ProjectId)
          {
              // Envelope capacities map (can be pulled from DB if needed)
              var envCaps = await _context.EnvelopesTypes
                  .Select(e => new { e.EnvelopeName, e.Capacity })
                  .ToListAsync();
              Dictionary<string, int> envelopeCapacities = envCaps
                  .ToDictionary(x => x.EnvelopeName, x => x.Capacity);

              var nrData = await _context.NRDatas
                  .Where(p => p.ProjectId == ProjectId)
                  .ToListAsync();

              var envBreaking = await _context.EnvelopeBreakages
                  .Where(p => p.ProjectId == ProjectId)
                  .ToListAsync();

              var extras = await _context.ExtrasEnvelope
                  .Where(p => p.ProjectId == ProjectId)
                  .ToListAsync();

              var projectconfig = await _context.ProjectConfigs
                  .Where(p => p.ProjectId == ProjectId)
                  .Select(p => p.Envelope)
                  .ToListAsync();

              var extrasconfig = await _context.ExtraConfigurations
                  .Where(p => p.ProjectId == ProjectId)
                  .ToListAsync();


              // Build dictionary for fast lookup of TotalEnv by NrDataId
              var envDict = envBreaking.ToDictionary(e => e.NrDataId, e => e.TotalEnvelope);

              var resultList = new List<object>();
              int serialnumber = 0;
              string prevNodalCode = null;
              string prevCatchNo = null;
              string prevMergeField = null; // CatchNo + CenterCode
              string prevExtraMergeField = null;
              int centerEnvCounter = 0;
              int extraCenterEnvCounter = 0;
              var nodalExtrasAddedForCatchNo = new HashSet<string>();
              var catchExtrasAdded = new HashSet<(int ExtraId, string CatchNo)>();

              // Extract outer envelope config from project config
              int nrOuterCapacity = 100; // default
              string outerEnvJson = projectconfig.FirstOrDefault();
              Console.WriteLine(outerEnvJson);
              if (!string.IsNullOrEmpty(outerEnvJson))
              {
                  var outerEnvDict = JsonSerializer.Deserialize<Dictionary<string, string>>(outerEnvJson);
                  Console.WriteLine(outerEnvDict);

                  // Explicitly check for the "Outer" key
                  var outerKey = outerEnvDict.FirstOrDefault(kvp => kvp.Key == "Outer").Key;
                  Console.WriteLine(outerKey);

                  if (outerKey != null && envelopeCapacities.ContainsKey(outerEnvDict[outerKey]))
                  {
                      nrOuterCapacity = envelopeCapacities[outerEnvDict[outerKey]];
                      Console.WriteLine(nrOuterCapacity);
                  }
              }

              // Helper method to add extra envelopes
              void AddExtraWithEnv(ExtraEnvelopes extra, string examDate, string examTime, string course, string subject, int NrQuantity)
              {
                  var extraConfig = extrasconfig.FirstOrDefault(e => e.ExtraType == extra.ExtraId);
                  int envCapacity = 100; // default fallback

                  if (extraConfig != null && !string.IsNullOrEmpty(extraConfig.EnvelopeType))
                  {
                      var envType = JsonSerializer.Deserialize<Dictionary<string, string>>(extraConfig.EnvelopeType);
                      if (envType.TryGetValue("Outer", out string outerType))
                      {
                          if (envelopeCapacities.TryGetValue(outerType, out int cap))
                              envCapacity = cap;
                      }
                  }

                  int totalEnv = (int)Math.Ceiling((double)extra.Quantity / envCapacity);
                  int quantityLeft = extra.Quantity;
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
                          int filledCapacity = envCapacity * (totalEnv - 1);
                          envQuantity = extra.Quantity - filledCapacity;
                      }
                      else
                      {
                          envQuantity = Math.Min(quantityLeft, envCapacity);
                      }
                      extraCenterEnvCounter++;

                      resultList.Add(new
                      {

                          SerialNumber = serialnumber++,
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
                          CourseName = course,
                          SubjectName = subject,

                          TotalEnv = totalEnv,
                          Env = $"{j}/{totalEnv}",
                          NRQuantity = NrQuantity,
                      });

                      quantityLeft -= envQuantity;
                  }
              }

              bool serialResetFlag = false;
              for (int i = 0; i < nrData.Count; i++)
              {
                  var current = nrData[i];
                  if (prevCatchNo != null && current.CatchNo != prevCatchNo)
                  {
                      // Reset serial number if CatchNo changes
                      serialnumber = 1;  // Start the serial number from 1 again
                      serialResetFlag = true; // Indicate that serial number was reset
                  }
                  // ➕ Nodal Extra (1) when NodalCode changes
                  if (prevNodalCode != null && current.NodalCode != prevNodalCode)
                  {
                      if (!nodalExtrasAddedForCatchNo.Contains(prevCatchNo))
                      {
                          var extrasToAdd = extras.Where(e => e.ExtraId == 1 && e.CatchNo == prevCatchNo).ToList();
                          foreach (var extra in extrasToAdd)
                          {
                              AddExtraWithEnv(extra,current.ExamDate,current.ExamTime,current.SubjectName,current.CourseName, current.NRQuantity);
                          }
                          nodalExtrasAddedForCatchNo.Add(prevCatchNo);
                      }
                  }

                  // ➕ Catch Extras (2, 3) when CatchNo changes
                  if (prevCatchNo != null && current.CatchNo != prevCatchNo)
                  {
                      foreach (var extraId in new[] { 2, 3 })
                      {
                          if (!catchExtrasAdded.Contains((extraId, prevCatchNo)))
                          {
                              var extrasToAdd = extras.Where(e => e.ExtraId == extraId && e.CatchNo == prevCatchNo).ToList();
                              foreach (var extra in extrasToAdd)
                              {
                                  AddExtraWithEnv(extra, current.ExamDate, current.ExamTime, current.SubjectName, current.CourseName,current.NRQuantity);
                              }
                              catchExtrasAdded.Add((extraId, prevCatchNo));
                          }
                      }
                  }

                  // ➕ Add current NRData row with TotalEnv replication
                  envDict.TryGetValue(current.Id, out int totalEnv);
                  if (totalEnv <= 0) totalEnv = 1;

                  int quantityLeft = current?.Quantity ?? 0;
                  Console.WriteLine(quantityLeft);
                  // Calculate CenterEnv based on merge field (CatchNo + CenterCode)
                  string currentMergeField = $"{current.CatchNo}-{current.CenterCode}";
                  Console.WriteLine(currentMergeField);
                  if (currentMergeField != prevMergeField)
                  {
                      centerEnvCounter = 0; // Reset counter for new merge group
                      Console.WriteLine(centerEnvCounter);
                      prevMergeField = currentMergeField;
                  }
                  else
                  {
                      centerEnvCounter++; // Increment for same merge group
                      Console.WriteLine(centerEnvCounter);
                  }

                  for (int j = 1; j <= totalEnv; j++)
                  {
                      int envQuantity;
                      centerEnvCounter++;
                      if (current.Quantity > nrOuterCapacity && j == 1)
                      {

                          envQuantity = current.Quantity - (totalEnv - 1) * nrOuterCapacity;
                          Console.WriteLine(envQuantity);
                      }
                      else if (current.Quantity <= nrOuterCapacity)
                      {
                          // Total quantity fits in one envelope
                          envQuantity = current?.Quantity??0;
                      }
                      else
                      {
                          // Subsequent envelopes get full capacity
                          envQuantity = nrOuterCapacity;
                          Console.WriteLine(envQuantity);
                      }

                      resultList.Add(new
                      {
                          SerialNumber = serialnumber++,
                          current.CourseName,
                          current.SubjectName,
                          current.CatchNo,
                          current.CenterCode,
                          current.ExamTime,
                          current.ExamDate,
                          current.Quantity,
                          EnvQuantity = envQuantity,
                          current.NodalCode,
                          current.NRDatas,
                          CenterEnv = centerEnvCounter, // Sequential number within same CatchNo+CenterCode
                          TotalEnv = totalEnv,
                          Env = $"{centerEnvCounter}/{totalEnv}", // Using CenterEnv for display
                          current.NRQuantity,
                      });

                      quantityLeft -= envQuantity;
                  }

                  prevNodalCode = current.NodalCode;
                  prevCatchNo = current.CatchNo;
              }

              // 🔁 Final extras for the last CatchNo and NodalCode
              if (prevCatchNo != null)
              {
                  if (!nodalExtrasAddedForCatchNo.Contains(prevCatchNo))
                  {
                      var extrasToAdd = extras.Where(e => e.ExtraId == 1 && e.CatchNo == prevCatchNo).ToList();
                      foreach (var extra in extrasToAdd)
                      {
                          var current = nrData.FirstOrDefault(p => p.CatchNo == prevCatchNo);

                          if (current != null)
                          {
                              // Pass the required parameters explicitly
                              AddExtraWithEnv(extra, current.ExamDate, current.ExamTime, current.SubjectName, current.CourseName, current.NRQuantity);
                          }

                      }
                  }

                  foreach (var extraId in new[] { 2, 3 })
                  {
                      if (!catchExtrasAdded.Contains((extraId, prevCatchNo)))
                      {
                          var extrasToAdd = extras.Where(e => e.ExtraId == extraId && e.CatchNo == prevCatchNo).ToList();
                          foreach (var extra in extrasToAdd)
                          {
                              var current = nrData.FirstOrDefault(p => p.CatchNo == prevCatchNo);
                              if (current != null)
                              {
                                  // Pass the required parameters explicitly
                                  AddExtraWithEnv(extra, current.ExamDate, current.ExamTime, current.SubjectName, current.CourseName, current.NRQuantity);
                              }
                          }
                      }
                  }
              }

              // Generate Excel Report
              using (var package = new ExcelPackage())
              {
                  var worksheet = package.Workbook.Worksheets.Add("BreakingResult");

                  // Add headers
                  worksheet.Cells[1, 1].Value = "Serial Number";
                  worksheet.Cells[1, 2].Value = "Course Name";
                  worksheet.Cells[1, 3].Value = "Subject Name";
                  worksheet.Cells[1, 4].Value = "Catch No";
                  worksheet.Cells[1, 5].Value = "Center Code";
                  worksheet.Cells[1, 6].Value = "Exam Time";
                  worksheet.Cells[1, 7].Value = "Exam Date";
                  worksheet.Cells[1, 8].Value = "Quantity";
                  worksheet.Cells[1, 9].Value = "EnvQuantity";
                  worksheet.Cells[1, 10].Value = "Nodal Code";
                  worksheet.Cells[1, 11].Value = "NR Datas";
                  worksheet.Cells[1, 12].Value = "Center Env";
                  worksheet.Cells[1, 13].Value = "Total Env";
                  worksheet.Cells[1, 14].Value = "Env";
                  worksheet.Cells[1, 15].Value = "NRQuantity";

                  // Style headers
                  using (var range = worksheet.Cells[1, 1, 1, 18])
                  {
                      range.Style.Font.Bold = true;
                      range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                      range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                      range.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                  }

                  // Add data rows
                  int row = 2;
                  foreach (var item in resultList)
                  {
                      var itemType = item.GetType();

                      // Use reflection to get property values
                      var serialNumber = itemType.GetProperty("SerialNumber")?.GetValue(item);
                      var courseName = itemType.GetProperty("CourseName")?.GetValue(item);
                      var subjectName = itemType.GetProperty("SubjectName")?.GetValue(item);
                      var catchNo = itemType.GetProperty("CatchNo")?.GetValue(item);
                      var centerCode = itemType.GetProperty("CenterCode")?.GetValue(item);
                      var examTime = itemType.GetProperty("ExamTime")?.GetValue(item);
                      var examDate = itemType.GetProperty("ExamDate")?.GetValue(item);
                      var quantity = itemType.GetProperty("Quantity")?.GetValue(item);
                      var envquantity = itemType.GetProperty("EnvQuantity")?.GetValue(item);
                      var nodalCode = itemType.GetProperty("NodalCode")?.GetValue(item);
                      var nrDatas = itemType.GetProperty("NRDatas")?.GetValue(item);
                      var centerEnv = itemType.GetProperty("CenterEnv")?.GetValue(item);
                      var totalEnv = itemType.GetProperty("TotalEnv")?.GetValue(item);
                      var env = itemType.GetProperty("Env")?.GetValue(item);
                      var nrQty = itemType.GetProperty("NRQuantity")?.GetValue(item);

                      worksheet.Cells[row, 1].Value = serialNumber;
                      worksheet.Cells[row, 2].Value = courseName?.ToString();
                      worksheet.Cells[row, 3].Value = subjectName?.ToString();
                      worksheet.Cells[row, 4].Value = catchNo?.ToString();
                      worksheet.Cells[row, 5].Value = centerCode?.ToString();
                      worksheet.Cells[row, 6].Value = examTime?.ToString();
                      worksheet.Cells[row, 7].Value = examDate;
                      worksheet.Cells[row, 8].Value = quantity;
                      worksheet.Cells[row, 9].Value = envquantity;
                      worksheet.Cells[row, 10].Value = nodalCode?.ToString();
                      worksheet.Cells[row, 11].Value = nrDatas?.ToString();
                      worksheet.Cells[row, 12].Value = centerEnv;
                      worksheet.Cells[row, 13].Value = totalEnv;
                      worksheet.Cells[row, 14].Value = env?.ToString();
                      worksheet.Cells[row, 15].Value = nrQty;
                      row++;
                  }

                  // Auto-fit columns
                  worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();

                  // Save the file
                  var reportPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", ProjectId.ToString());
                  if (!Directory.Exists(reportPath))
                  {
                      Directory.CreateDirectory(reportPath);
                  }

                  var fileName = $"BreakingReport.xlsx";
                  var filePath = Path.Combine(reportPath, fileName);
                  if (System.IO.File.Exists(filePath))
                  {
                      System.IO.File.Delete(filePath);
                  }
                  var fileInfo = new FileInfo(filePath);
                  package.SaveAs(fileInfo);

                  // Return both the data and the file path
                  return Ok(new
                  {
                      Result = resultList,
                  });
              }
          }*/

        // DELETE: api/EnvelopeBreakages/5
        [HttpDelete("{id}")]
        public async Task<IActionResult> DeleteEnvelopeBreakage(int id)
        {
            var envelopeBreakage = await _context.EnvelopeBreakages.FindAsync(id);
            if (envelopeBreakage == null)
            {
                return NotFound();
            }

            _context.EnvelopeBreakages.Remove(envelopeBreakage);
            await _context.SaveChangesAsync();

            return NoContent();
        }

        private bool EnvelopeBreakageExists(int id)
        {
            return _context.EnvelopeBreakages.Any(e => e.EnvelopeId == id);
        }
    }
}
