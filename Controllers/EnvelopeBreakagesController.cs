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
                    ["Quantity"] = row.Quantity
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
                foreach (var size in outerSizes)
                {
                    int count = remaining / size;
                    if (count > 0)
                    {
                        outerBreakdown[$"E{size}"] = count.ToString();
                        remaining -= count * size;
                    }
                }

                var envelope = new EnvelopeBreakage
                {
                    ProjectId = ProjectId,
                    NrDataId = row.Id,
                    InnerEnvelope = JsonSerializer.Serialize(innerBreakdown),
                    OuterEnvelope = JsonSerializer.Serialize(outerBreakdown)
                };

                // ✅ Add to database
                _context.EnvelopeBreakages.Add(envelope);
            }

            await _context.SaveChangesAsync();

            return Ok("Envelope breakdown successfully saved to EnvelopeBreakage table.");
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
