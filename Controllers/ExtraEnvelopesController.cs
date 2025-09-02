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
                .GroupBy(d => d.CatchNo)
                .Select(g => new
                {
                    CatchNo = g.Key,
                    Quantity = g.Sum(x => x.Quantity)
                })
                .ToListAsync();

            var extraConfig = await _context.ExtraConfigurations
                .Where(c => c.ProjectId == ProjectId).ToListAsync();

            if (extraConfig == null || !extraConfig.Any())
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

                foreach (var data in nrDataList)
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
                                calculatedQuantity = (int)Math.Ceiling((double)(data.Quantity * percentValue) / 100);
                            }
                            break;

                        case "Range":
                            // calculatedQuantity = HandleRangeLogic(data, extraConfig);
                            break;
                    }

                    int innerCount = (int)Math.Ceiling((double)calculatedQuantity / innerCapacity);
                    int outerCount = (int)Math.Ceiling((double)calculatedQuantity / outerCapacity);

                    var envelope = new ExtraEnvelopes
                    {
                        ProjectId = ProjectId,
                        CatchNo = data.CatchNo,
                        ExtraId = config.ExtraType,
                        Quantity = calculatedQuantity,
                        InnerEnvelope = innerCount.ToString(),
                        OuterEnvelope = outerCount.ToString(),
                    };

                    envelopesToAdd.Add(envelope);

                    // Log
                    Console.WriteLine($"CatchNo {data.CatchNo} → Quantity: {calculatedQuantity}, Inner: {innerCount} packets, Outer: {outerCount} packets");
                }
            }
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
                return BadRequest($"Duplicate ExtraEnvelopes detected for the following combinations: {duplicateList}");
            }

            await _context.ExtrasEnvelope.AddRangeAsync(envelopesToAdd);
            await _context.SaveChangesAsync();

            var reportPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "reports");
            Directory.CreateDirectory(reportPath);
            var filename = $"ExtraEnvelope {ProjectId}.xlsx";
            var filePath = Path.Combine(reportPath, filename);

            // Assuming reportRows is defined elsewhere or fetched here, otherwise add fetching logic
            var reportRows = await _context.NRDatas
                .Where(d => d.ProjectId == ProjectId)
                .ToListAsync();

            // Gather static properties (excluding NRDatas)
            var baseProperties = typeof(NRData).GetProperties(BindingFlags.Public | BindingFlags.Instance)
                                               .Where(p => p.Name != "NRDatas")
                                               .ToList();

            var extraHeaders = new HashSet<string>();

            // Group reportRows by Nodal property and assign ExtraTypeId incrementally
            var groupedByNodal = reportRows.GroupBy(r => r.NodalCode).ToList();

            var rowsWithExtraTypeId = new List<(NRData row, Dictionary<string, string> extras, int ExtraTypeId)>();

            int extraTypeCounter = 1;
            foreach (var group in groupedByNodal)
            {
                foreach (var row in group)
                {
                    Dictionary<string, string> extras = new();

                    if (!string.IsNullOrEmpty(row.NRDatas))
                    {
                        try
                        {
                            extras = JsonSerializer.Deserialize<Dictionary<string, string>>(row.NRDatas) ?? new();
                            foreach (var key in extras.Keys)
                                extraHeaders.Add(key);
                        }
                        catch
                        {
                            // Optionally log parsing error
                        }
                    }

                    rowsWithExtraTypeId.Add((row, extras, extraTypeCounter));
                }
                extraTypeCounter++;
            }

            using (var package = new ExcelPackage())
            {
                var ws = package.Workbook.Worksheets.Add("Merge Report");

                // Write Headers
                var properties = typeof(NRData).GetProperties(BindingFlags.Public | BindingFlags.Instance);
                int col = 1;
                foreach (var prop in properties)
                {
                    ws.Cells[1, col].Value = prop.Name;
                    ws.Cells[1, col].Style.Font.Bold = true;
                    col++;
                }

                int rowIdx = 2;

                // Group NRData by NodalCode

                foreach (var nodalGroup in groupedByNodal)
                {
                    // 1. Write all original NRData rows in this Nodal group
                    foreach (var data in nodalGroup)
                    {
                        col = 1;
                        foreach (var prop in properties)
                        {
                            var value = prop.GetValue(data);
                            ws.Cells[rowIdx, col++].Value = value?.ToString() ?? "";
                        }
                        rowIdx++;
                    }

                    // 2. Write ExtraId == 1 for this Nodal group (matching CatchNo)
                    var nodalCatchNos = nodalGroup.Select(x => x.CatchNo).ToHashSet();

                    var extrasType1 = envelopesToAdd
                        .Where(e => e.ExtraId == 1 && nodalCatchNos.Contains(e.CatchNo))
                        .ToList();

                    foreach (var extra in extrasType1)
                    {
                        // Find base NRData row
                        var baseData = reportRows.FirstOrDefault(r => r.CatchNo == extra.CatchNo);
                        if (baseData == null) continue;

                        col = 1;
                        foreach (var prop in properties)
                        {
                            object? value = null;
                            switch (prop.Name)
                            {
                                case nameof(NRData.Quantity):
                                    value = extra.Quantity;
                                    break;

                                case nameof(NRData.CenterCode):
                                    value = "Nodal Extra";
                                    break;

                                default:
                                    value = prop.GetValue(baseData);
                                    break;
                            }

                            ws.Cells[rowIdx, col++].Value = value?.ToString() ?? "";
                        }
                        rowIdx++;
                    }
                }

                // 3. Write ExtraId == 2 and 3 rows (after all nodals)
                var extrasType2 = envelopesToAdd
                    .Where(e => e.ExtraId == 2)
                    .ToList();

                foreach (var extra in extrasType2)
                {
                    var baseData = reportRows.FirstOrDefault(r => r.CatchNo == extra.CatchNo);
                    if (baseData == null) continue;

                    col = 1;
                    foreach (var prop in properties)
                    {
                        object? value = null;
                        switch (prop.Name)
                        {
                            case nameof(NRData.Quantity):
                                value = extra.Quantity;
                                break;

                            case nameof(NRData.CenterCode):
                                value = "University Extra";
                                break;

                            case nameof(NRData.NodalCode):
                                value = "Extras";
                                    break;

                            default:
                                value = prop.GetValue(baseData);
                                break;
                        }

                        ws.Cells[rowIdx, col++].Value = value?.ToString() ?? "";
                    }
                    rowIdx++;
                }

                var extrasType3 = envelopesToAdd
                   .Where(e => e.ExtraId == 3)
                   .ToList();

                foreach (var extra in extrasType3)
                {
                    var baseData = reportRows.FirstOrDefault(r => r.CatchNo == extra.CatchNo);
                    if (baseData == null) continue;

                    col = 1;
                    foreach (var prop in properties)
                    {
                        object? value = null;
                        switch (prop.Name)
                        {
                            case nameof(NRData.Quantity):
                                value = extra.Quantity;
                                break;

                            case nameof(NRData.CenterCode):
                                value = "Office Extra";
                                break;

                            case nameof(NRData.NodalCode):
                                value = "Extras";
                                break;

                            default:
                                value = prop.GetValue(baseData);
                                break;
                        }

                        ws.Cells[rowIdx, col++].Value = value?.ToString() ?? "";
                    }
                    rowIdx++;
                }
                ws.Cells[ws.Dimension.Address].AutoFitColumns();
                package.SaveAs(new FileInfo(filePath));
            }


            return Ok(envelopesToAdd);
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
