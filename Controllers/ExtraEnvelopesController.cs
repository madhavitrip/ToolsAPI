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
using Tools.Migrations;

namespace Tools.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ExtraEnvelopesController : ControllerBase
    {
        private readonly ERPToolsDbContext _context;

        public ExtraEnvelopesController(ERPToolsDbContext context)
        {
            _context = context;
        }

        // GET: api/ExtraEnvelopes
        [HttpGet]
        public async Task<ActionResult<IEnumerable<ExtraEnvelopes>>> GetExtrasEnvelope(int ProjectId)
        {
            var NrData = await _context.NRDatas.Where(p=>p.ProjectId == ProjectId).ToListAsync();

            return await _context.ExtrasEnvelope.ToListAsync();
        }

        // GET: api/ExtraEnvelopes/5
        [HttpGet("{id}")]
        public async Task<ActionResult<ExtraEnvelopes>> GetExtraEnvelopes(int id)
        {
            var extraEnvelopes = await _context.ExtrasEnvelope.FindAsync(id);

            if (extraEnvelopes == null)
            {
                return NotFound();
            }

            return extraEnvelopes;
        }

        // PUT: api/ExtraEnvelopes/5
        // To protect from overposting attacks, see https://go.microsoft.com/fwlink/?linkid=2123754
        [HttpPut("{id}")]
        public async Task<IActionResult> PutExtraEnvelopes(int id, ExtraEnvelopes extraEnvelopes)
        {
            if (id != extraEnvelopes.Id)
            {
                return BadRequest();
            }

            _context.Entry(extraEnvelopes).State = EntityState.Modified;

            try
            {
                await _context.SaveChangesAsync();
            }
            catch (DbUpdateConcurrencyException)
            {
                if (!ExtraEnvelopesExists(id))
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

        public class EnvelopeType
        {
            public string Inner { get; set; }
            public string Outer { get; set; }
        }


        // POST: api/ExtraEnvelopes
        // To protect from overposting attacks, see https://go.microsoft.com/fwlink/?linkid=2123754
        [HttpPost]
        public async Task<ActionResult> PostExtraEnvelopes(int ProjectId)
        {
            var nrDataList = await _context.NRDatas
                .Where(d => d.ProjectId == ProjectId)
                .ToListAsync();

            var groupedNR = nrDataList
                .GroupBy(x => x.CatchNo)
                .Select(g => new
                {
                    CatchNo = g.Key,
                    Quantity = g.Sum(x => x.Quantity ?? 0)
                })
                .ToList();

            var extraConfig = await _context.ExtraConfigurations
                .Where(c => c.ProjectId == ProjectId)
                .ToListAsync();

            if (!extraConfig.Any())
                return BadRequest("No ExtraConfiguration found for the project.");

            var envelopesToAdd = new List<ExtraEnvelopes>();
            foreach (var config in extraConfig)
            {
                EnvelopeType envelopeType;
                try
                {
                    envelopeType = JsonSerializer.Deserialize<EnvelopeType>(config.EnvelopeType);
                }
                catch
                {
                    return BadRequest($"Invalid EnvelopeType JSON for ExtraType {config.ExtraType}");
                }

                int innerCapacity = GetEnvelopeCapacity(envelopeType.Inner);
                int outerCapacity = GetEnvelopeCapacity(envelopeType.Outer);

                foreach (var data in groupedNR)
                {
                    int calculatedQuantity = 0;
                    switch (config.Mode)
                    {
                        case "Fixed":
                            calculatedQuantity = int.Parse(config.Value);
                            break;
                        case "Percentage":
                            if (decimal.TryParse(config.Value, out var percentValue))
                                calculatedQuantity = (int)Math.Ceiling((double)(data.Quantity * percentValue) / 100);
                            break;
                    }

                    int innerCount = (int)Math.Ceiling((double)calculatedQuantity / innerCapacity);
                    int outerCount = (int)Math.Ceiling((double)calculatedQuantity / outerCapacity);

                    envelopesToAdd.Add(new ExtraEnvelopes
                    {
                        ProjectId = ProjectId,
                        CatchNo = data.CatchNo,
                        ExtraId = config.ExtraType,
                        Quantity = calculatedQuantity,
                        InnerEnvelope = innerCount.ToString(),
                        OuterEnvelope = outerCount.ToString(),
                    });
                }
            }

            // ❌ Check for duplicates
            var existingCombinations = await _context.ExtrasEnvelope
                .Where(e => e.ProjectId == ProjectId)
                .Select(e => new { e.ProjectId, e.CatchNo, e.ExtraId })
                .ToListAsync();

            var duplicates = envelopesToAdd
                .Where(e => existingCombinations.Any(existing =>
                    existing.ProjectId == e.ProjectId &&
                    existing.CatchNo == e.CatchNo &&
                    existing.ExtraId == e.ExtraId))
                .ToList();

            if (duplicates.Any())
            {
                var duplicateList = string.Join(", ", duplicates.Select(d => $"CatchNo: {d.CatchNo}, ExtraId: {d.ExtraId}"));
                return BadRequest($"Duplicate ExtraEnvelopes detected for: {duplicateList}");
            }

            await _context.ExtrasEnvelope.AddRangeAsync(envelopesToAdd);
            await _context.SaveChangesAsync();

            // ------------------- 📊 Generate Excel Report -------------------
            var allNRData = await _context.NRDatas
                .Where(x => x.ProjectId == ProjectId)
                .ToListAsync();

            var envelopeBreakages = await _context.EnvelopeBreakages
                .Where(x => x.ProjectId == ProjectId)
                .ToListAsync();

            var extraconfig = await _context.ExtraConfigurations
                .Where(x => x.ProjectId == ProjectId) .ToListAsync();

            var envelopeDict = envelopeBreakages.ToDictionary(e => e.NrDataId);

            var groupedByNodal = allNRData.GroupBy(x => x.NodalCode).ToList();

            var extraHeaders = new HashSet<string>();
            var innerKeys = new HashSet<string>();
            var outerKeys = new HashSet<string>();

            var allRows = new List<Dictionary<string, object>>();

            foreach (var nodalGroup in groupedByNodal)
            {
                foreach (var row in nodalGroup)
                {
                    var dict = NRDataToDictionary(row, envelopeDict, extraHeaders, innerKeys, outerKeys);
                    allRows.Add(dict);
                }

                // ➕ Add ExtraTypeId = 1 for this nodal group
                var catchNos = nodalGroup.Select(x => x.CatchNo).ToHashSet();
                var extras1 = envelopesToAdd.Where(e => e.ExtraId == 1 && catchNos.Contains(e.CatchNo)).ToList();

                foreach (var extra in extras1)
                {
                    var baseRow = allNRData.FirstOrDefault(x => x.CatchNo == extra.CatchNo);
                    if (baseRow == null) continue;

                    var dict = NRDataToDictionary(baseRow, envelopeDict, extraHeaders, innerKeys, outerKeys);
                    dict["Quantity"] = extra.Quantity;
                    dict["CenterCode"] = "Nodal Extra";
                    dict["InnerEnvelope"] = extra.InnerEnvelope;
                    dict["OuterEnvelope"] = extra.OuterEnvelope;
                    allRows.Add(dict);
                }
            }

            // ➕ Add ExtraTypeId = 2 (University) and 3 (Office) at the end
            foreach (var extraType in new[] { 2, 3 })
            {
                var extras = envelopesToAdd.Where(e => e.ExtraId == extraType).ToList();

                foreach (var extra in extras)
                {
                    var baseRow = allNRData.FirstOrDefault(x => x.CatchNo == extra.CatchNo);
                    if (baseRow == null) continue;

                    var dict = NRDataToDictionary(baseRow, envelopeDict, extraHeaders, innerKeys, outerKeys);
                    dict["Quantity"] = extra.Quantity;
                    dict["CenterCode"] = extraType == 2 ? "University Extra" : "Office Extra";
                    dict["NodalCode"] = "Extras";
                    dict["InnerEnvelope"] = extra.InnerEnvelope;
                    dict["OuterEnvelope"] = extra.OuterEnvelope;
                    allRows.Add(dict);
                }
            }

            // Create Excel
            var allHeaders = typeof(Tools.Models.NRData).GetProperties().Select(p => p.Name).ToList();
            allHeaders.AddRange(extraHeaders.OrderBy(x => x));
            allHeaders.AddRange(innerKeys.OrderBy(x => x));
            allHeaders.AddRange(outerKeys.OrderBy(x => x));

            var filePath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "reports", $"ExtraEnvelope_{ProjectId}.xlsx");
            Directory.CreateDirectory(Path.GetDirectoryName(filePath)!);

            using (var package = new ExcelPackage())
            {
                var ws = package.Workbook.Worksheets.Add("Extra Envelope");

                for (int i = 0; i < allHeaders.Count; i++)
                {
                    ws.Cells[1, i + 1].Value = allHeaders[i];
                    ws.Cells[1, i + 1].Style.Font.Bold = true;
                }

                int rowIdx = 2;
                foreach (var row in allRows)
                {
                    for (int colIdx = 0; colIdx < allHeaders.Count; colIdx++)
                    {
                        var key = allHeaders[colIdx];
                        row.TryGetValue(key, out object value);
                        ws.Cells[rowIdx, colIdx + 1].Value = value?.ToString() ?? "";
                    }
                    rowIdx++;
                }

                ws.Cells[ws.Dimension.Address].AutoFitColumns();
                package.SaveAs(new FileInfo(filePath));
            }

            return Ok(envelopesToAdd);
        }

        private Dictionary<string, object> NRDataToDictionary(
        Tools.Models.NRData row,
            Dictionary<int, Tools.Models.EnvelopeBreakage> envMap,
            HashSet<string> extraKeys,
            HashSet<string> innerKeys,
            HashSet<string> outerKeys)
        {
            var dict = new Dictionary<string, object>
            {
                ["Id"] = row.Id,
                ["ProjectId"] = row.ProjectId,
                ["CourseName"] = row.CourseName,
                ["SubjectName"] = row.SubjectName,
                ["CatchNo"] = row.CatchNo,
                ["CenterCode"] = row.CenterCode,
                ["ExamTime"] = row.ExamTime,
                ["ExamDate"] = row.ExamDate,
                ["Quantity"] = row.Quantity,
                ["NodalCode"] = row.NodalCode,
                ["NRDatas"] = row.NRDatas
            };

            if (!string.IsNullOrEmpty(row.NRDatas))
            {
                try
                {
                    var extras = JsonSerializer.Deserialize<Dictionary<string, string>>(row.NRDatas);
                    if (extras != null)
                    {
                        foreach (var kvp in extras)
                        {
                            dict[kvp.Key] = kvp.Value;
                            extraKeys.Add(kvp.Key);
                        }
                    }
                }
                catch { }
            }

            if (envMap.TryGetValue(row.Id, out var env))
            {
                void ParseEnv(string? json, HashSet<string> keySet, string prefix)
                {
                    if (!string.IsNullOrEmpty(json))
                    {
                        try
                        {
                            var envDict = JsonSerializer.Deserialize<Dictionary<string, string>>(json);
                            if (envDict != null)
                            {
                                foreach (var kvp in envDict)
                                {
                                    string key = $"{prefix}_{kvp.Key}";
                                    dict[key] = kvp.Value;
                                    keySet.Add(key);
                                }
                            }
                        }
                        catch { }
                    }
                }

                ParseEnv(env.InnerEnvelope, innerKeys, "Inner");
                ParseEnv(env.OuterEnvelope, outerKeys, "Outer");
            }

            return dict;
        }


        private int GetEnvelopeCapacity(string envelopeCode)
        {
            if (string.IsNullOrWhiteSpace(envelopeCode))
                return 1; // default to 1 if null or invalid

            // Expecting format like "E10", "E25", etc.
            var numberPart = new string(envelopeCode.Where(char.IsDigit).ToArray());

            return int.TryParse(numberPart, out var capacity) ? capacity : 1;
        }


        // DELETE: api/ExtraEnvelopes/5
        [HttpDelete("{id}")]
        public async Task<IActionResult> DeleteExtraEnvelopes(int id)
        {
            var extraEnvelopes = await _context.ExtrasEnvelope.FindAsync(id);
            if (extraEnvelopes == null)
            {
                return NotFound();
            }

            _context.ExtrasEnvelope.Remove(extraEnvelopes);
            await _context.SaveChangesAsync();

            return NoContent();
        }

        private bool ExtraEnvelopesExists(int id)
        {
            return _context.ExtrasEnvelope.Any(e => e.Id == id);
        }
    }
}
