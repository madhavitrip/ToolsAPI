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

namespace Tools.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ExtraEnvelopesController : ControllerBase
    {
        private readonly ERPToolsDbContext _context;
        private readonly ILoggerService _loggerService;

        public ExtraEnvelopesController(ERPToolsDbContext context, ILoggerService loggerService)
        {
            _context = context;
            _loggerService = loggerService;
        }

        // GET: api/ExtraEnvelopes
        [HttpGet]
        public async Task<ActionResult<IEnumerable<ExtraEnvelopes>>> GetExtrasEnvelope(int ProjectId)
        {
            var NrData = await _context.NRDatas.Where(p => p.ProjectId == ProjectId && p.Status == true).ToListAsync();

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
                _loggerService.LogEvent($"Updated ExtraEnvelope with id {id}", "ExtraEnvelopes", User.Identity?.Name != null ? int.Parse(User.Identity.Name) : 0, extraEnvelopes.ProjectId);
                await _context.SaveChangesAsync();
            }
            catch (Exception ex)
            {
                if (!ExtraEnvelopesExists(id))
                {
                    _loggerService.LogEvent($"ExtraEnvelopes with ID {id} not found during updating", "ExtraEnvelopes", User.Identity?.Name != null ? int.Parse(User.Identity.Name) : 0, extraEnvelopes.ProjectId);
                    return NotFound();
                }
                else
                {
                    _loggerService.LogError("Error updating ExtraEnvelopes", ex.Message, nameof(ExtraEnvelopesController));
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
            try
            {
                // ✅ Get GroupId
                var project = await _context.Projects
                    .Where(p => p.ProjectId == ProjectId)
                    .Select(p => new { p.GroupId })
                    .FirstOrDefaultAsync();

                if (project == null)
                    return BadRequest("Project not found");

                bool isOnlyReport = project.GroupId == 28;

                List<ExtraEnvelopes> envelopesToUse = new();

                // =====================================================
                // ✅ CASE 1: REPORT ONLY (GroupId = 28)
                // =====================================================
                if (isOnlyReport)
                {
                    envelopesToUse = await _context.ExtrasEnvelope
                        .Where(e => e.ProjectId == ProjectId)
                        .ToListAsync();

                    if (!envelopesToUse.Any())
                        return BadRequest("No existing ExtraEnvelope data found for report.");
                }
                else
                {
                    // =====================================================
                    // ✅ CASE 2: NORMAL FLOW (CALCULATE + SAVE)
                    // =====================================================

                    var nrDataList = await _context.NRDatas
                        .Where(d => d.ProjectId == ProjectId && d.Status == true)
                        .ToListAsync();

                    var groupedNR = nrDataList
                        .GroupBy(x => x.CatchNo)
                        .Select(g => new
                        {
                            CatchNo = g.Key,
                            Quantity = g.Sum(x => x.Quantity)
                        })
                        .ToList();

                    if (!groupedNR.Any())
                        return BadRequest("No grouping found");

                    var extraConfig = await _context.ExtraConfigurations
                        .Where(c => c.ProjectId == ProjectId)
                        .ToListAsync();

                    if (!extraConfig.Any())
                        return BadRequest("No ExtraConfiguration found");

                    var existingCombinations = await _context.ExtrasEnvelope
                        .Where(e => e.ProjectId == ProjectId)
                        .ToListAsync();

                    if (existingCombinations.Any())
                    {
                        _context.ExtrasEnvelope.RemoveRange(existingCombinations);
                        await _context.SaveChangesAsync();
                    }

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
                                    {
                                        var rawQuantity = (double)(data.Quantity * percentValue) / 100;

                                        if (innerCapacity > 10)
                                            calculatedQuantity = (int)Math.Ceiling(rawQuantity / innerCapacity) * innerCapacity;
                                        else
                                            calculatedQuantity = (int)Math.Ceiling(rawQuantity / outerCapacity) * outerCapacity;
                                    }
                                    break;

                                case "Range":
                                    if (!string.IsNullOrEmpty(config.RangeConfig))
                                    {
                                        var rangeConfig = JsonSerializer.Deserialize<RangeConfigModel>(config.RangeConfig);

                                        var range = rangeConfig?.ranges?
                                            .FirstOrDefault(r => data.Quantity >= r.from && data.Quantity <= r.to);

                                        if (range != null)
                                            calculatedQuantity = range.value;
                                    }
                                    break;
                            }

                            int innerCount = innerCapacity > 0
                                ? (int)Math.Ceiling((double)calculatedQuantity / innerCapacity)
                                : 0;

                            int outerCount = outerCapacity > 0
                                ? (int)Math.Ceiling((double)calculatedQuantity / outerCapacity)
                                : 0;

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

                    await _context.ExtrasEnvelope.AddRangeAsync(envelopesToAdd);
                    await _context.SaveChangesAsync();

                    envelopesToUse = envelopesToAdd;
                }

                // =====================================================
                // ✅ COMMON EXCEL GENERATION (USED BY BOTH MODES)
                // =====================================================

                var allNRData = await _context.NRDatas
                    .Where(x => x.ProjectId == ProjectId)
                    .ToListAsync();

                var extraConfigList = await _context.ExtraConfigurations
                    .Where(x => x.ProjectId == ProjectId)
                    .ToListAsync();

                var groupedByNodal = allNRData.GroupBy(x => x.NodalCode).ToList();

                var extraHeaders = new HashSet<string>();
                var innerKeys = new HashSet<string>();
                var outerKeys = new HashSet<string>();

                var allRows = new List<Dictionary<string, object>>();

                foreach (var nodalGroup in groupedByNodal)
                {
                    var catchNos = nodalGroup.Select(x => x.CatchNo).ToHashSet();

                    var extras1 = envelopesToUse
                        .Where(e => e.ExtraId == 1 && catchNos.Contains(e.CatchNo))
                        .ToList();

                    foreach (var extra in extras1)
                    {
                        var baseRow = allNRData.FirstOrDefault(x => x.CatchNo == extra.CatchNo);
                        var config = extraConfigList.FirstOrDefault(c => c.ExtraType == extra.ExtraId);

                        if (baseRow == null || config == null) continue;

                        var dict = ExtraEnvelopeToDictionary(baseRow, extra, config, extraHeaders, innerKeys, outerKeys);
                        allRows.Add(dict);
                    }
                }

                foreach (var extraType in new[] { 2, 3 })
                {
                    var extras = envelopesToUse.Where(e => e.ExtraId == extraType).ToList();

                    foreach (var extra in extras)
                    {
                        var baseRow = allNRData.FirstOrDefault(x => x.CatchNo == extra.CatchNo);
                        var config = extraConfigList.FirstOrDefault(c => c.ExtraType == extra.ExtraId);

                        if (baseRow == null || config == null) continue;

                        var dict = ExtraEnvelopeToDictionary(baseRow, extra, config, extraHeaders, innerKeys, outerKeys);
                        allRows.Add(dict);
                    }
                }

                if (!allRows.Any())
                    return BadRequest("No valid data for Excel");

                var headers = typeof(NRData).GetProperties()
                    .Select(p => p.Name)
                    .Where(p => p != "Id" && p != "ProjectId")
                    .ToList();

                headers.AddRange(extraHeaders);
                headers.AddRange(innerKeys);
                headers.AddRange(outerKeys);

                var path = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", ProjectId.ToString());
                if (!Directory.Exists(path))
                    Directory.CreateDirectory(path);

                var filePath = Path.Combine(path, "ExtrasCalculation.xlsx");

                if (System.IO.File.Exists(filePath))
                    System.IO.File.Delete(filePath);

                using (var package = new ExcelPackage())
                {
                    var ws = package.Workbook.Worksheets.Add("Extra Envelope");

                    for (int i = 0; i < headers.Count; i++)
                        ws.Cells[1, i + 1].Value = headers[i];

                    int row = 2;

                    foreach (var data in allRows)
                    {
                        for (int col = 0; col < headers.Count; col++)
                        {
                            data.TryGetValue(headers[col], out object val);
                            ws.Cells[row, col + 1].Value = val?.ToString();
                        }
                        row++;
                    }

                    ws.Cells[ws.Dimension.Address].AutoFitColumns();
                    package.SaveAs(new FileInfo(filePath));
                }

                return Ok(new
                {
                    message = isOnlyReport
                        ? "Report generated from existing data (GroupId = 28)"
                        : "Data calculated, saved, and report generated",
                    data = envelopesToUse
                });
            }
            catch (Exception ex)
            {
                _loggerService.LogError("Error creating ExtraEnvelope", ex.Message, nameof(ExtraEnvelopesController));
                return StatusCode(500, "Internal server error");
            }
        }
        public class RangeConfigModel
        {
            public string rangeType { get; set; }
            public List<RangeItem> ranges { get; set; }
        }

        public class RangeItem
        {
            public int from { get; set; }
            public int to { get; set; }
            public int value { get; set; }
        }
        private Dictionary<string, object> ExtraEnvelopeToDictionary(
       Tools.Models.NRData baseRow,
       ExtraEnvelopes extra,
       ExtrasConfiguration config,
       HashSet<string> extraKeys,
       HashSet<string> innerKeys,
       HashSet<string> outerKeys)
        {
            var dict = new Dictionary<string, object>
            {
                ["CourseName"] = baseRow.CourseName,
                ["SubjectName"] = baseRow.SubjectName,
                ["CatchNo"] = baseRow.CatchNo,
                ["ExamTime"] = baseRow.ExamTime,
                ["ExamDate"] = baseRow.ExamDate,
                ["NodalCode"] = "",
                ["Quantity"] = extra.Quantity,
                ["CenterSort"] = baseRow.CenterCode,
                ["RouteSort"] = baseRow.RouteSort
            };

            // Set CenterCode
            switch (extra.ExtraId)
            {
                case 1:
                    dict["CenterCode"] = "Nodal Extra";
                        dict["NodalSort"] =baseRow.NodalSort;
                        dict["CenterSort"] = 10000;
                    dict["RouteSort"] = baseRow.RouteSort;
                    break;
                case 2:
                    dict["CenterCode"] = "University Extra";
                    dict["NodalSort"] = 10000;
                    dict["CenterSort"] = 100000;
                    dict["RouteSort"] = 10000;
                    break;
                case 3:
                    dict["CenterCode"] = "Office Extra";
                    dict["NodalSort"] = 100000;
                    dict["CenterSort"] = 1000000;
                    dict["RouteSort"] = 100000;
                    break;
                default:
                    dict["CenterCode"] = "Extra";
                    break;
            }

            // Add NRDatas (extra fields)
            if (!string.IsNullOrEmpty(baseRow.NRDatas))
            {
                try
                {
                    var extras = JsonSerializer.Deserialize<Dictionary<string, string>>(baseRow.NRDatas);
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

            // ✅ Set InnerEnvelope and OuterEnvelope values in proper columns
            EnvelopeType envelopeType=null;
            try
            {
                envelopeType = JsonSerializer.Deserialize<EnvelopeType>(config.EnvelopeType);
            }
            catch
            {
            }

            string innerKey = null, outerKey = null;

            if (envelopeType != null)
            {
                if (!string.IsNullOrWhiteSpace(envelopeType.Inner) && !string.IsNullOrWhiteSpace(extra.InnerEnvelope))
                {
                    innerKey = $"Inner_{envelopeType.Inner}";
                    dict[innerKey] = extra.InnerEnvelope;
                    innerKeys.Add(innerKey);
                }

                if (!string.IsNullOrWhiteSpace(envelopeType.Outer) && !string.IsNullOrWhiteSpace(extra.OuterEnvelope))
                {
                    outerKey = $"Outer_{envelopeType.Outer}";
                    dict[outerKey] = extra.OuterEnvelope;
                    outerKeys.Add(outerKey);
                }
            }

            return dict;
        }

        private int GetEnvelopeCapacity(string envelopeCode)
        {
            if (string.IsNullOrWhiteSpace(envelopeCode))
                return 0; // default to 1 if null or invalid

            // Expecting format like "E10", "E25", etc.
            var numberPart = new string(envelopeCode.Where(char.IsDigit).ToArray());

            return int.TryParse(numberPart, out var capacity) ? capacity : 1;
        }


        // DELETE: api/ExtraEnvelopes/5
        [HttpDelete("{id}")]
        public async Task<IActionResult> DeleteExtraEnvelopes(int id)
        {
            try
            {
                var extraEnvelopes = await _context.ExtrasEnvelope.FindAsync(id);
                if (extraEnvelopes == null)
                {
                    return NotFound();
                }

                _context.ExtrasEnvelope.Remove(extraEnvelopes);
                await _context.SaveChangesAsync();
                _loggerService.LogEvent($"ExtraEnvelope with ID {id} is deleted", "ExtraEnvelope", User.Identity?.Name != null ? int.Parse(User.Identity.Name) : 0, extraEnvelopes.ProjectId);
                return NoContent();
            }
            catch (Exception ex)
            {
                _loggerService.LogError($"Error deleting ExtraEnvelope with ID {id}", ex.Message, nameof(ExtraEnvelopesController));
                return StatusCode(500, "Internal Server Error");
            }
        }

        private bool ExtraEnvelopesExists(int id)
        {
            return _context.ExtrasEnvelope.Any(e => e.Id == id);
        }
    }
}
