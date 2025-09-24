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

        /*   [HttpGet("Replication")]
           public async Task<IActionResult> ReplicationConfiguration(int ProjectId, string SortingFields)
           {
               var nrData = await _context.NRDatas
                   .Where(p => p.ProjectId == ProjectId)
                   .ToListAsync();

               var envBreaking = await _context.EnvelopeBreakages
                   .Where(p => p.ProjectId == ProjectId)
                   .ToListAsync();

               var extras = await _context.ExtrasEnvelope
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

               string prevNodalCode = null;
               string prevCatchNo = null;

               // Track which extras have been added to avoid duplicates
               var nodalExtrasAddedForCatchNo = new HashSet<string>(); // CatchNos for which nodal extra (ExtraId=1) was added
               var catchExtrasAdded = new HashSet<(int ExtraId, string CatchNo)>(); // For ExtraId 2 & 3

               for (int i = 0; i < nrData.Count; i++)
               {
                   var current = nrData[i];

                   // When NodalCode changes (and not first item), add ExtraId=1 for the CatchNo of previous row ONLY
                   if (prevNodalCode != null && current.NodalCode != prevNodalCode)
                   {
                       if (!nodalExtrasAddedForCatchNo.Contains(prevCatchNo))
                       {
                           var extrasToAdd = extras.Where(e => e.ExtraId == 1 && e.CatchNo == prevCatchNo).ToList();

                           foreach (var extra in extrasToAdd)
                           {
                               resultList.Add(new
                               {
                                   ExtraAttached = true,
                                   ExtraId = 1,
                                   extra.CatchNo,
                                   extra.Quantity,
                                   extra.InnerEnvelope,
                                   extra.OuterEnvelope,
                                   CenterCode = "Nodal Extra"
                               });
                           }
                           nodalExtrasAddedForCatchNo.Add(prevCatchNo);
                       }
                   }

                   // When CatchNo changes (and not first item), add ExtraId=2 and 3 for the previous CatchNo ONLY
                   if (prevCatchNo != null && current.CatchNo != prevCatchNo)
                   {
                       foreach (var extraId in new[] { 2, 3 })
                       {
                           if (!catchExtrasAdded.Contains((extraId, prevCatchNo)))
                           {
                               var extrasToAdd = extras.Where(e => e.ExtraId == extraId && e.CatchNo == prevCatchNo).ToList();
                               foreach (var extra in extrasToAdd)
                               {
                                   resultList.Add(new
                                   {
                                       ExtraAttached = true,
                                       extra.ExtraId,
                                       extra.CatchNo,
                                       extra.Quantity,
                                       extra.InnerEnvelope,
                                       extra.OuterEnvelope,
                                       CenterCode = extraId == 2 ? "University Extra" : "Office Extra"
                                   });
                               }
                               catchExtrasAdded.Add((extraId, prevCatchNo));
                           }
                       }
                   }

                   // Add the current NRData with TotalEnv attached
                   envDict.TryGetValue(current.Id, out int totalEnv);

                   resultList.Add(new
                   {
                       current.Id,
                       current.CourseName,
                       current.SubjectName,
                       current.CatchNo,
                       current.CenterCode,
                       current.ExamTime,
                       current.ExamDate,
                       current.Quantity,
                       current.NodalCode,
                       current.NRDatas,
                       TotalEnv = totalEnv
                   });

                   prevNodalCode = current.NodalCode;
                   prevCatchNo = current.CatchNo;
               }

               // After the loop, add extras for the last CatchNo and NodalCode group if not already added
               if (prevNodalCode != null && prevCatchNo != null)
               {
                   if (!nodalExtrasAddedForCatchNo.Contains(prevCatchNo))
                   {
                       var extrasToAdd = extras.Where(e => e.ExtraId == 1 && e.CatchNo == prevCatchNo).ToList();
                       foreach (var extra in extrasToAdd)
                       {
                           resultList.Add(new
                           {
                               ExtraAttached = true,
                               ExtraId = 1,
                               extra.CatchNo,
                               extra.Quantity,
                               extra.InnerEnvelope,
                               extra.OuterEnvelope,
                               CenterCode = "Nodal Extra"
                           });
                       }
                       nodalExtrasAddedForCatchNo.Add(prevCatchNo);
                   }

                   foreach (var extraId in new[] { 2, 3 })
                   {
                       if (!catchExtrasAdded.Contains((extraId, prevCatchNo)))
                       {
                           var extrasToAdd = extras.Where(e => e.ExtraId == extraId && e.CatchNo == prevCatchNo).ToList();
                           foreach (var extra in extrasToAdd)
                           {
                               resultList.Add(new
                               {
                                   ExtraAttached = true,
                                   extra.ExtraId,
                                   extra.CatchNo,
                                   extra.Quantity,
                                   extra.InnerEnvelope,
                                   extra.OuterEnvelope,
                                   CenterCode = extraId == 2 ? "University Extra" : "Office Extra"
                               });
                           }
                           catchExtrasAdded.Add((extraId, prevCatchNo));
                       }
                   }
               }

               return Ok(new
               {
                   Result = resultList
               });
           }
   */

        [HttpGet("Replication")]
        public async Task<IActionResult> ReplicationConfiguration(int ProjectId, string SortingFields)
        {
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
                .Where(p => p.ProjectId == ProjectId).Select(p => p.Envelope).ToListAsync();
            var extrasconfig = await _context.ExtraConfigurations
                .Where(p => p.ProjectId == ProjectId).ToListAsync();

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

            string prevNodalCode = null;
            string prevCatchNo = null;

            var nodalExtrasAddedForCatchNo = new HashSet<string>();
            var catchExtrasAdded = new HashSet<(int ExtraId, string CatchNo)>();

            void AddExtraWithEnv(ExtraEnvelopes extra)
            {
                if (int.TryParse(extra.OuterEnvelope, out int outerCount) && outerCount > 0)
                {
                    for (int j = 1; j <= outerCount; j++)
                    {
                        resultList.Add(new
                        {
                            ExtraAttached = true,
                            extra.ExtraId,
                            extra.CatchNo,
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
                            Env = $"{j}/{outerCount}"
                        });
                    }
                }
            }

            // Function to parse envelope sizes from JSON string
            int GetEnvelopeSizeFromConfig(string envelopeJson, string key)
            {
                var envelopeConfig = JsonSerializer.Deserialize<Dictionary<string, string>>(envelopeJson);
                if (envelopeConfig != null && envelopeConfig.ContainsKey(key))
                {
                    return int.TryParse(envelopeConfig[key], out var size) ? size : 0;
                }
                return 0;
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

                // ➕ Add current NRData row with TotalEnv replication and Quantity splitting
                envDict.TryGetValue(current.Id, out int totalEnv);
                if (totalEnv <= 0) totalEnv = 1; // fallback

                // Retrieve the envelope size from EnvelopeBreakages or ExtraConfig
                var envConfig = envBreaking.FirstOrDefault(e => e.ProjectId == ProjectId);
                int envelopeSize = 0;
                if (envConfig != null)
                {
                    // Assuming we are using "E100" for the primary envelope size from Env config.
                    envelopeSize = GetEnvelopeSizeFromConfig(envConfig.OuterEnvelope, "E100");
                }

                // If no envelope size found, fallback to extras configuration
                if (envelopeSize == 0)
                {
                    var extraConfig = extrasconfig.FirstOrDefault(e => e.ProjectId == ProjectId && e.ExtraType == 2); // For example, using ExtraType 2
                    if (extraConfig != null)
                    {
                        envelopeSize = GetEnvelopeSizeFromConfig(extraConfig.Value, "Outer"); // Extract size from outer envelope for extras
                    }
                }

                int totalQuantity = current?.Quantity??0;
                int remainingQuantity = totalQuantity;
                int envCount = totalEnv;

                // Distribute quantity across envelopes
                for (int j = 1; j <= envCount; j++)
                {
                    int currentQuantity = envelopeSize;

                    // Adjust the quantity for the last envelope
                    if (remainingQuantity < envelopeSize && remainingQuantity > 0)
                    {
                        currentQuantity = remainingQuantity;
                    }

                    // Create result for the envelope with the adjusted quantity
                    resultList.Add(new
                    {
                        current.Id,
                        current.CourseName,
                        current.SubjectName,
                        current.CatchNo,
                        current.CenterCode,
                        current.ExamTime,
                        current.ExamDate,
                        Quantity = currentQuantity, // Use the adjusted quantity
                        current.NodalCode,
                        current.NRDatas,
                        TotalEnv = totalEnv,
                        Env = $"{j}/{totalEnv}"
                    });

                    // Decrease the remaining quantity for the next envelope
                    remainingQuantity -= currentQuantity;
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

            return Ok(new
            {
                Result = resultList
            });
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
