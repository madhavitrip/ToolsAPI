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

            var reportPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "reports");
            Directory.CreateDirectory(reportPath);

            var filename = $"EnvelopeBreaking_{ProjectId}.xlsx";
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

            // 🧾 Generate Excel report
            using (var package = new ExcelPackage())
            {
                var ws = package.Workbook.Worksheets.Add("Envelope Report");

                // Write headers
                for (int i = 0; i < allHeaders.Count; i++)
                {
                    ws.Cells[1, i + 1].Value = allHeaders[i];
                    ws.Cells[1, i + 1].Style.Font.Bold = true;
                }

                // Write data rows
                int rowIdx = 2;
                foreach (var rowDict in parsedRows)
                {
                    for (int colIdx = 0; colIdx < allHeaders.Count; colIdx++)
                    {
                        var key = allHeaders[colIdx];
                        rowDict.TryGetValue(key, out object value);
                        ws.Cells[rowIdx, colIdx + 1].Value = value?.ToString() ?? "";
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

            return Ok("Envelope breakdown successfully saved to EnvelopeBreakage table.");
        }


        [HttpGet("Replication")]
        public async Task<IActionResult> ReplicationConfiguration(int ProjectId, string SortingFields)
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

            var extrasconfig = await _context.ExtraConfigurations
                .Where(p => p.ProjectId == ProjectId)
                .ToListAsync();

            // Handle sorting
            if (!string.IsNullOrEmpty(SortingFields))
            {
                var sortFields = SortingFields.Split('-');
                IOrderedEnumerable<NRData> orderedData = null;

                foreach (var (field, index) in sortFields.Select((value, i) => (value, i)))
                {
                    var property = typeof(NRData).GetProperty(field);
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
                        Quantity = envQuantity,
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
                        Quantity = envQuantity,
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
            var reportPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "reports");

            // Ensure the directory exists
            Directory.CreateDirectory(reportPath);

            var filename = $"EnvelopeMaking_{ProjectId}.xlsx";
            var filePath = Path.Combine(reportPath, filename);

            // 📁 Skip generation if file already exists
            if (System.IO.File.Exists(filePath))
            {
                return Ok(new { message = "File already exists", filePath }); // Still return data for UI
            }

            // Generate Excel file and save it to the directory
            using (var package = new ExcelPackage())
            {
                // Create a worksheet
                var worksheet = package.Workbook.Worksheets.Add("ReplicationResult");

                // Add headers to the worksheet
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
                // Fill in data
                int row = 2; // Start from row 2 (because row 1 is for headers)
                int serialNumber = 1;
                string preCatchNo = null;
                int prevEnd = 0;

                foreach (var item in resultList)
                {
                    var catchNo = item.GetType().GetProperty("CatchNo")?.GetValue(item)?.ToString();
                    var centerCode = item.GetType().GetProperty("CenterCode")?.GetValue(item)?.ToString();
                    var quantity = Convert.ToInt32(item.GetType().GetProperty("Quantity")?.GetValue(item));
                    var totalEnv = Convert.ToInt32(item.GetType().GetProperty("TotalEnv")?.GetValue(item));

                    // Calculate the Start column based on the CatchNo
                    int start = (catchNo != prevCatchNo) ? 1 : prevEnd + 1;
                    int end = start + totalEnv - 1; // End is Start + TotalEnv - 1
                    string serial = $"{start} to {end}";

                    worksheet.Cells[row, 1].Value = item.GetType().GetProperty("Serialnumber")?.GetValue(item);
                    worksheet.Cells[row, 2].Value = catchNo;
                    worksheet.Cells[row, 3].Value = centerCode;
                    worksheet.Cells[row, 4].Value = item.GetType().GetProperty("ExamTime")?.GetValue(item);
                    worksheet.Cells[row, 5].Value = item.GetType().GetProperty("ExamDate")?.GetValue(item);
                    worksheet.Cells[row, 6].Value = quantity;
                    worksheet.Cells[row, 7].Value = item.GetType().GetProperty("NodalCode")?.GetValue(item);
                    worksheet.Cells[row, 8].Value =totalEnv;
                    worksheet.Cells[row, 9].Value = item.GetType().GetProperty("Env")?.GetValue(item);
                    worksheet.Cells[row, 10].Value = start; // Start
                    worksheet.Cells[row, 11].Value = end; // End
                    worksheet.Cells[row, 12].Value = serial; // Serial (Start to End)

                    preCatchNo = catchNo;
                    prevEnd = end;
                    row++;
                }

                // Save the Excel package to the specified path
                FileInfo fi = new FileInfo(filePath);
                package.SaveAs(fi);
            }

            // Return success message or file path
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
