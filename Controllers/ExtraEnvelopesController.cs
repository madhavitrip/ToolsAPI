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

            return await _context.ExtrasEnvelope
                .Where(e => e.ProjectId == ProjectId && e.Status == 1)
                .ToListAsync();
        }

        public class ExtraEnvelopeByCatchRequest
        {
            public int ProjectId { get; set; }
            public List<string> CatchNos { get; set; } = new();
        }

        [HttpPost("ByCatchNos")]
        public async Task<ActionResult<IEnumerable<ExtraEnvelopes>>> GetExtrasByCatchNos(
            [FromBody] ExtraEnvelopeByCatchRequest request)
        {
            if (request == null || request.ProjectId <= 0)
            {
                return BadRequest("ProjectId is required.");
            }

            var normalizedCatchNos = (request.CatchNos ?? new List<string>())
                .Select(c => (c ?? string.Empty).Trim())
                .Where(c => !string.IsNullOrWhiteSpace(c))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            if (!normalizedCatchNos.Any())
            {
                return Ok(new List<ExtraEnvelopes>());
            }

            var extras = await _context.ExtrasEnvelope
                .Where(e =>
                    e.ProjectId == request.ProjectId &&
                    e.Status == 1 &&
                    e.CatchNo != null &&
                    normalizedCatchNos.Contains(e.CatchNo))
                .ToListAsync();

            var uniqueExtras = extras
                .GroupBy(e => new { CatchNo = e.CatchNo, ExtraId = e.ExtraId })
                .Select(group =>
                {
                    var latest = group.OrderByDescending(e => e.Id).First();
                    return new ExtraEnvelopes
                    {
                        Id = latest.Id,
                        ProjectId = latest.ProjectId,
                        CatchNo = latest.CatchNo,
                        ExtraId = latest.ExtraId,
                        Quantity = latest.Quantity,
                        InnerEnvelope = latest.InnerEnvelope,
                        OuterEnvelope = latest.OuterEnvelope,
                        Status = latest.Status,
                    };
                })
                .ToList();

            return Ok(uniqueExtras);
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
                await _loggerService.LogEventAsync($"Updated ExtraEnvelope with id {id}", "ExtraEnvelopes", LogHelper.GetTriggeredBy(User), extraEnvelopes.ProjectId);
                await _context.SaveChangesAsync();
            }
            catch (Exception ex)
            {
                if (!ExtraEnvelopesExists(id))
                {
                    await _loggerService.LogEventAsync($"ExtraEnvelopes with ID {id} not found during updating", "ExtraEnvelopes", LogHelper.GetTriggeredBy(User), extraEnvelopes.ProjectId);
                    return NotFound();
                }
                else
                {
                    await _loggerService.LogErrorAsync("Error saving extra envelopes", ex.Message, nameof(ExtraEnvelopesController));
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
        public async Task<ActionResult> PostExtraEnvelopes(int ProjectId, int? uploadId = null)
        {
            try
            {
                var projectConfig = await _context.ProjectConfigs
                    .FirstOrDefaultAsync(p => p.ProjectId == ProjectId);

                var eligibleSteps = Tools.Models.PipelineNavigator.GetEligiblePickupSteps(Tools.Models.PipelineNavigator.STEP_AWAITING_EXTRA);

                List<NRData> nrDataList;
                if (uploadId.HasValue)
                {
                    var all = await _context.NRDatas.Where(p => p.ProjectId == ProjectId).ToListAsync();
                    nrDataList = all.Where(x => x.UploadList != null && x.UploadList.Contains(uploadId.Value)).ToList();
                }
                else
                {
                    nrDataList = await _context.NRDatas
                        .Where(d => d.ProjectId == ProjectId && d.Status == true && eligibleSteps.Contains(d.Steps))
                        .ToListAsync();
                }

                if (!nrDataList.Any())
                    return BadRequest("No suitable data found for Extra Configuration. Either no data is at the required step or all data is already processed.");

                var project = await _context.Projects
                    .Where(p => p.ProjectId == ProjectId)
                    .Select(p => new { p.GroupId })
                    .FirstOrDefaultAsync();

                if (project == null)
                    return BadRequest("Project not found");

                bool isOnlyReport = project.GroupId == 8;

                List<ExtraEnvelopes> envelopesToUse = new();

                // ================= REPORT ONLY =================
                if (isOnlyReport)
                {
                    envelopesToUse = await _context.ExtrasEnvelope
                        .Where(e => e.ProjectId == ProjectId && e.Status == 1)
                        .ToListAsync();

                    if (!envelopesToUse.Any())
                        return BadRequest("No existing ExtraEnvelope data found for report.");
                }
                else
                {
                    var extraConfig = await _context.ExtraConfigurations
                        .Where(c => c.ProjectId == ProjectId)
                        .ToListAsync();

                    if (!extraConfig.Any())
                        return BadRequest("No ExtraConfiguration found");

                    // Remove old
                    var existing = await _context.ExtrasEnvelope
                        .Where(e => e.ProjectId == ProjectId)
                        .ToListAsync();

                    if (existing.Any())
                    {
                        _context.ExtrasEnvelope.RemoveRange(existing);
                        await _context.SaveChangesAsync();
                    }

                    var envelopesToAdd = new List<ExtraEnvelopes>();

                    foreach (var config in extraConfig)
                    {
                        bool useNodal = !string.IsNullOrWhiteSpace(config.nodalValue);

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

                        // ---------- GROUPING ----------
                        var groupedData = useNodal
                            ? nrDataList
                                .GroupBy(x => new { x.CatchNo, NodalCode = x.NodalCode ?? "" })
                                .Select(g => new
                                {
                                    g.Key.CatchNo,
                                    g.Key.NodalCode,
                                    Quantity = g.Sum(x => x.Quantity)
                                }).ToList()
                            : nrDataList
                                .GroupBy(x => x.CatchNo)
                                .Select(g => new
                                {
                                    CatchNo = g.Key,
                                    NodalCode = (string)null,
                                    Quantity = g.Sum(x => x.Quantity)
                                }).ToList();

                        // ---------- NODAL CONFIG ----------
                        List<NodalValueConfig> nodalConfigs = null;
                        if (useNodal)
                        {
                            try
                            {
                                nodalConfigs = JsonSerializer.Deserialize<List<NodalValueConfig>>(
                                    config.nodalValue,
                                    new JsonSerializerOptions { PropertyNameCaseInsensitive = true });
                            }
                            catch { }
                        }

                        foreach (var data in groupedData)
                        {
                            int calculatedQuantity = 0;

                            if (useNodal && nodalConfigs != null)
                            {
                                var match = nodalConfigs.FirstOrDefault(nc =>
                                    !string.IsNullOrEmpty(nc.NodalCodes) &&
                                    nc.NodalCodes.Split(',')
                                        .Select(n => n.Trim())
                                        .Contains(data.NodalCode, StringComparer.OrdinalIgnoreCase));

                                if (match != null)
                                {
                                    switch (config.Mode)
                                    {
                                        case "Fixed":
                                            if (int.TryParse(match.Value, out var fv))
                                                calculatedQuantity = fv;
                                            break;

                                        case "Percentage":
                                            if (decimal.TryParse(match.Value, out var pv))
                                                calculatedQuantity = (int)Math.Round((double)(data.Quantity * pv) / 100);
                                            break;
                                    }
                                }
                            }
                            else
                            {
                                switch (config.Mode)
                                {
                                    case "Fixed":
                                        if (int.TryParse(config.Value, out var fqv))
                                            calculatedQuantity = fqv;
                                        break;

                                    case "Percentage":
                                        if (decimal.TryParse(config.Value, out var percent))
                                        {
                                            var raw = (double)(data.Quantity * percent) / 100;

                                            if (innerCapacity > 10)
                                                calculatedQuantity = (int)Math.Ceiling(raw / innerCapacity) * innerCapacity;
                                            else if (outerCapacity > 0)
                                                calculatedQuantity = (int)Math.Ceiling(raw / outerCapacity) * outerCapacity;
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
                                NodalCode = data.NodalCode,
                                ExtraId = config.ExtraType,
                                Quantity = calculatedQuantity,
                                InnerEnvelope = innerCount.ToString(),
                                OuterEnvelope = outerCount.ToString(),
                                Status = 1
                            });
                        }
                    }

                    await _context.ExtrasEnvelope.AddRangeAsync(envelopesToAdd);

                    foreach (var nr in nrDataList)
                        nr.Steps = Tools.Models.PipelineNavigator.GetNextStep(Tools.Models.PipelineNavigator.STEP_AWAITING_EXTRA, projectConfig?.Modules);

                    await _context.SaveChangesAsync();

                    envelopesToUse = envelopesToAdd;
                }

                // ================= EXCEL =================
                var extraConfigs = await _context.ExtraConfigurations
                    .Where(x => x.ProjectId == ProjectId)
                    .ToListAsync();

                var extraHeaders = new HashSet<string>();
                var innerKeys = new HashSet<string>();
                var outerKeys = new HashSet<string>();

                var allRows = new List<Dictionary<string, object>>();

                foreach (var extra in envelopesToUse)
                {
                    var baseRow = nrDataList.FirstOrDefault(x =>
                        x.CatchNo == extra.CatchNo &&
                        (extra.NodalCode == null || x.NodalCode == extra.NodalCode));

                    var config = extraConfigs.FirstOrDefault(c => c.ExtraType == extra.ExtraId);

                    if (baseRow == null || config == null) continue;

                    var dict = ExtraEnvelopeToDictionary(baseRow, extra, config, extraHeaders, innerKeys, outerKeys);
                    allRows.Add(dict);
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
                if (!Directory.Exists(path)) Directory.CreateDirectory(path);

                var fileName = uploadId.HasValue ? $"ExtrasCalculation_v{uploadId}.xlsx" : "ExtrasCalculation.xlsx";
                var filePath = Path.Combine(path, fileName);
                if (System.IO.File.Exists(filePath)) System.IO.File.Delete(filePath);

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
                        ? "Report generated from existing data"
                        : "Calculated and report generated",
                    data = envelopesToUse,
                    fileName = fileName
                });
            }
            catch (Exception ex)
            {
                await _loggerService.LogErrorAsync("Error creating ExtraEnvelope", ex.Message, nameof(ExtraEnvelopesController));
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

        public class NodalValueConfig
        {
            public string NodalCodes { get; set; }
            public string Value { get; set; }
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

            // ? Set InnerEnvelope and OuterEnvelope values in proper columns
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
                await _loggerService.LogEventAsync("Extra Envelopes configuration saved", "ExtraEnvelopes", LogHelper.GetTriggeredBy(User), extraEnvelopes.ProjectId);
                return NoContent();
            }
            catch (Exception ex)
            {
                await _loggerService.LogErrorAsync($"Error deleting ExtraEnvelope with ID {id}", ex.Message, nameof(ExtraEnvelopesController));
                return StatusCode(500, "Internal Server Error");
            }
        }

        private bool ExtraEnvelopesExists(int id)
        {
            return _context.ExtrasEnvelope.Any(e => e.Id == id);
        }
    }
}

