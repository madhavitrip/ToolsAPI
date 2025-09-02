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
                int col = 1;
                foreach (var prop in baseProperties)
                {
                    ws.Cells[1, col].Value = prop.Name;
                    ws.Cells[1, col].Style.Font.Bold = true;
                    col++;
                }

                // Add ExtraTypeId header before extraHeaders
                ws.Cells[1, col].Value = "ExtraTypeId";
                ws.Cells[1, col].Style.Font.Bold = true;
                col++;

                var extraHeaderList = extraHeaders.OrderBy(k => k).ToList();
                foreach (var key in extraHeaderList)
                {
                    ws.Cells[1, col].Value = key;
                    ws.Cells[1, col].Style.Font.Bold = true;
                    col++;
                }

                // Write Rows
                int rowIdx = 2;
                foreach (var (item, extras, extraTypeId) in rowsWithExtraTypeId)
                {
                    col = 1;

                    foreach (var prop in baseProperties)
                    {
                        var value = prop.GetValue(item);
                        ws.Cells[rowIdx, col++].Value = value?.ToString() ?? "";
                    }

                    // Write ExtraTypeId
                    ws.Cells[rowIdx, col++].Value = extraTypeId;

                    foreach (var key in extraHeaderList)
                    {
                        extras.TryGetValue(key, out var val);
                        ws.Cells[rowIdx, col++].Value = val ?? "";
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
