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

            foreach (var row in nrDataList)
            {
                bool exists = await _context.EnvelopeBreakages
              .AnyAsync(e => e.NrDataId == row.Id);

                if (exists)
                    continue; // Skip adding this entry because it already exists

                int quantity = row.Quantity ?? 0;
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
            var response = await client.GetAsync($"https://localhost:7276/api/EnvelopeBreakages?ProjectId={ProjectId}");

            if (!response.IsSuccessStatusCode)
            {
                // Handle failure from GET call as needed
                return StatusCode((int)response.StatusCode, "Failed to get envelope breakages after configuration.");
            }

            var jsonString = await response.Content.ReadAsStringAsync();


            return Ok("Envelope breakdown successfully saved to EnvelopeBreakage table.");
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

            var envelopeIds = await _context.ProjectConfigs
           .Where(p => p.ProjectId == ProjectId)
           .Select(p => p.EnvelopeMakingCriteria)
           .FirstOrDefaultAsync();  // Assuming Envelope is a list or collection of IDs.

            var boxIds = await _context.ProjectConfigs
                .Where (p => p.ProjectId == ProjectId)
                .Select (p => p.BoxBreakingCriteria) .FirstOrDefaultAsync();

            var fields = await _context.Fields
                .Where(f=>boxIds.Contains(f.FieldId))
                .ToListAsync();

            // Step 2: Fetch the corresponding field names from the Fields table
            var fieldNames = await _context.Fields
                .Where(f => envelopeIds.Contains(f.FieldId))  // Assuming envelopeIds contains the IDs of the fields to sort by
                .Select(f => f.Name)  // Get the field names
                .ToListAsync();


            var extrasconfig = await _context.ExtraConfigurations
                .Where(p => p.ProjectId == ProjectId)
                .ToListAsync();

            // Handle sorting
            if (fieldNames.Any())
            {
                IOrderedEnumerable<NRData> orderedData = null;

                foreach (var (fieldName, index) in fieldNames.Select((value, i) => (value, i)))
                {
                    var property = typeof(NRData).GetProperty(fieldName);
                    if (property == null) continue;

                    if (index == 0)
                        orderedData = nrData.OrderBy(x => property.GetValue(x, null));
                    else
                        orderedData = orderedData.ThenBy(x => property.GetValue(x, null));
                }

                if (orderedData != null)
                    nrData = orderedData.ToList();
            }

            // Build dictionary for fast lookup of TotalEnv by NrDataId
            var envDict = envBreaking.ToDictionary(e => e.NrDataId, e => e.TotalEnvelope);

            var resultList = new List<object>();
            int serialnumber = 1;
            string prevNodalCode = null;
            string prevCatchNo = null;

            var nodalExtrasAddedForCatchNo = new HashSet<string>();
            var catchExtrasAdded = new HashSet<(int ExtraId, string CatchNo)>();

            // Extract outer envelope config from project config
            int nrOuterCapacity = 100; // default
            string outerEnvJson = projectconfig.FirstOrDefault();

            if (!string.IsNullOrEmpty(outerEnvJson))
            {
                var outerEnvDict = JsonSerializer.Deserialize<Dictionary<string, string>>(outerEnvJson);
                var outerKey = outerEnvDict.Keys.FirstOrDefault();
                if (outerKey != null && envelopeCapacities.ContainsKey(outerKey))
                {
                    nrOuterCapacity = envelopeCapacities[outerKey];
                }
            }

            // Helper method to add extra envelopes
            void AddExtraWithEnv(ExtraEnvelopes extra)
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

                    resultList.Add(new
                    {
                        Serialnumber = serialnumber++,
                        ExtraAttached = true,
                        extra.ExtraId,
                        extra.CatchNo,
                        EnvQuantity = envQuantity,
                        extra.Quantity,
                        extra.InnerEnvelope,
                        extra.OuterEnvelope,
                        CenterCode = extra.ExtraId switch
                        {
                            1 => "Nodal Extra",
                            2 => "University Extra",
                            3 => "Office Extra",
                            _ => "Extra"
                        },
                        Env = $"{j}/{totalEnv}"
                    });

                    quantityLeft -= envQuantity;
                }
            }

            for (int i = 0; i < nrData.Count; i++)
            {
                var current = nrData[i];

                // ➕ Nodal Extra (1) when NodalCode changes
                if (prevNodalCode != null && current.NodalCode != prevNodalCode)
                {
                    if (!nodalExtrasAddedForCatchNo.Contains(prevCatchNo))
                    {
                        var extrasToAdd = extras.Where(e => e.ExtraId == 1 && e.CatchNo == prevCatchNo).ToList();
                        foreach (var extra in extrasToAdd)
                        {
                            AddExtraWithEnv(extra);
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
                                AddExtraWithEnv(extra);
                            }
                            catchExtrasAdded.Add((extraId, prevCatchNo));
                        }
                    }
                }

                // ➕ Add current NRData row with TotalEnv replication
                envDict.TryGetValue(current.Id, out int totalEnv);
                if (totalEnv <= 0) totalEnv = 1;

                int quantityLeft = current?.Quantity ?? 0;

                for (int j = 1; j <= totalEnv; j++)
                {
                    int envQuantity;

                    if (j == 1 && totalEnv > 1)
                    {
                        int filledCapacity = nrOuterCapacity * (totalEnv - 1);
                        envQuantity = current?.Quantity - filledCapacity ?? 0;
                    }
                    else
                    {
                        envQuantity = Math.Min(quantityLeft, nrOuterCapacity);
                    }

                    resultList.Add(new
                    {
                        SerialNumber = serialnumber++,
                        current.CatchNo,
                        current.CenterCode,
                        current.ExamTime,
                        current.ExamDate,
                        EnvQuantity = envQuantity,
                        current.Quantity,
                        current.NodalCode,
                        TotalEnv = totalEnv,
                        Env = $"{j}/{totalEnv}"
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
                        AddExtraWithEnv(extra);
                    }
                }

                foreach (var extraId in new[] { 2, 3 })
                {
                    if (!catchExtrasAdded.Contains((extraId, prevCatchNo)))
                    {
                        var extrasToAdd = extras.Where(e => e.ExtraId == extraId && e.CatchNo == prevCatchNo).ToList();
                        foreach (var extra in extrasToAdd)
                        {
                            AddExtraWithEnv(extra);
                        }
                    }
                }
            }
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
                return Ok(new { message = "File already exists", filePath }); // Still return data for UI
            }

            // Step 1: Remove duplicates (CatchNo + CenterCode), preserving first occurrence
            var uniqueData = resultList
                .GroupBy(x =>
                    $"{x.GetType().GetProperty("CatchNo")?.GetValue(x)}_{x.GetType().GetProperty("CenterCode")?.GetValue(x)}"
                )
                .Select(g => g.First()) // Keep first
                .ToList();

            // Step 2: Calculate Start, End, Serial (before sorting)
            var enrichedList = new List<dynamic>();
            string previousCatchNo = null;
            int previousEnd = 0;

            foreach (var item in uniqueData)
            {
                var catchNo = item.GetType().GetProperty("CatchNo")?.GetValue(item)?.ToString();
                var centerCode = item.GetType().GetProperty("CenterCode")?.GetValue(item)?.ToString();
                var examTime = item.GetType().GetProperty("ExamTime")?.GetValue(item);
                var examDate = item.GetType().GetProperty("ExamDate")?.GetValue(item);
                var quantity = Convert.ToInt32(item.GetType().GetProperty("Quantity")?.GetValue(item));
                var nodalCode = item.GetType().GetProperty("NodalCode")?.GetValue(item)?.ToString();
                var totalEnv = Convert.ToInt32(item.GetType().GetProperty("TotalEnv")?.GetValue(item));
                var env = item.GetType().GetProperty("Env")?.GetValue(item)?.ToString();

                int start;
                if (catchNo != previousCatchNo)
                    start = 1;
                else
                    start = previousEnd + 1;

                int end = start + totalEnv - 1;
                string serial = $"{start} to {end}";

                enrichedList.Add(new
                {
                    CatchNo = catchNo,
                    CenterCode = centerCode,
                    ExamTime = examTime,
                    ExamDate = examDate,
                    Quantity = quantity,
                    NodalCode = nodalCode,
                    TotalEnv = totalEnv,
                    Env = env,
                    Start = start,
                    End = end,
                    Serial = serial
                });

                previousCatchNo = catchNo;
                previousEnd = end;
            }

            // Step 3: Assign SerialNumbers AFTER serial calculation, using original order
            int serialNumber = 1;
            var withSerials = enrichedList
                .Select(x => new
                {
                    SerialNumber = serialNumber++,
                    x.CatchNo,
                    x.CenterCode,
                    x.ExamTime,
                    x.ExamDate,
                    x.Quantity,
                    x.NodalCode,
                    x.TotalEnv,
                    x.Env,
                    x.Start,
                    x.End,
                    x.Serial
                })
                .ToList();

            // Step 4: Now SORT the data by CenterCode, then CatchNo
            var finalSorted = withSerials
                .OrderBy(x => x.CenterCode)
                .ThenBy(x => x.CatchNo)
                .ToList();

            // Step 6: Add TotalPages and BoxNo
            int boxNo = 1001;
            int runningPages = 0;
            string prevMergeKey = null;
            

            var finalWithBoxes = new List<dynamic>();

            foreach (var item in finalSorted)
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
                bool overflow = (runningPages + totalPages > 17000);

                if (mergeChanged || overflow)
                {
                    // case: overflow detected
                    if (overflow)
                    {
                        int leftover = (runningPages + totalPages) - 17000;
                        int filled = totalPages - leftover;

                        // if leftover < 50% threshold, split into two balanced boxes
                        if (leftover > 0 && leftover < 8500)
                        {
                            int combined = 17000 + leftover;
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
                                item.Env,
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
                                item.Env,
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
                    item.Env,
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
                worksheet.Cells[1, 9].Value = "Env";
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
                    worksheet.Cells[row, 9].Value = item.Env;
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
