using ERPToolsAPI.Data;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using System.Reflection;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;
using Tools.Models;
using System.Text.Json;
using Humanizer;

namespace Tools.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class DuplicateController : ControllerBase
    {
        private readonly ERPToolsDbContext _context;

        public DuplicateController(ERPToolsDbContext context)
        {
            _context = context;
        }

        [HttpPost]
        public async Task<IActionResult> MergeFields(int ProjectId, string mergefields, bool consolidate, bool enhancement, double Percent)
        {
            var mergeFields = mergefields.Split(',', StringSplitOptions.RemoveEmptyEntries)
                                          .Select(f => f.Trim().ToLower()).ToList();


            if (mergeFields.Count == 0)
                return BadRequest("No merge fields provided.");

            // Get project data
            var data = await _context.NRDatas
                .Where(p => p.ProjectId == ProjectId)
                .ToListAsync();

            if (!data.Any())
                return NotFound("No data found for this project.");

            // Group the data based on the merge fields
            var grouped = data.GroupBy(d =>
            {
                var key = new List<string>();
                foreach (var field in mergeFields)
                {
                    var value = d.GetType().GetProperty(field, BindingFlags.IgnoreCase | BindingFlags.Public | BindingFlags.Instance)
                                 ?.GetValue(d)?.ToString()?.Trim() ?? "";
                    key.Add(value);
                }
                return string.Join("|", key); // Composite key
            });

            int mergedCount = 0;
            var deletedRows = new List<NRData>();
            var reportRows = new List<NRData>();

            foreach (var group in grouped)
            {
                reportRows.AddRange(group); // Include all rows for reporting

                if (group.Count() <= 1)
                    continue;

                var keep = group.First();
                if (consolidate)
                {
                    keep.Quantity = group.Sum(x => x.Quantity);
                }


                var duplicates = group.Skip(1).ToList();
                _context.NRDatas.RemoveRange(duplicates);

                mergedCount += duplicates.Count;
                deletedRows.AddRange(duplicates);
            }


            if (enhancement)
            {
                foreach (var d in data)
                {
                    if (d.Quantity.HasValue)
                    {
                        d.Quantity = (int?)Math.Round(Percent * d.Quantity.Value);
                    }
                }
            }
            else
            {
                var innerEnv = await _context.ProjectConfigs.
                    Where(s => s.ProjectId == ProjectId).Select(s => s.Envelope)
                    .FirstOrDefaultAsync();
                if (!string.IsNullOrEmpty(innerEnv))
                {
                    try
                    {
                        var envelopeDict = JsonSerializer.Deserialize<Dictionary<string, string>>(innerEnv);
                        if (envelopeDict != null && envelopeDict.ContainsKey("Inner"))
                        {
                            var innerSizes = envelopeDict["Inner"]
                                .Split(',', StringSplitOptions.RemoveEmptyEntries)
                                .Select(e => e.Trim().ToUpper().Replace("E", "")) // get number from E10 → 10
                                .Where(x => int.TryParse(x, out _))
                                .Select(int.Parse)
                                .OrderBy(x => x)
                                .ToList();

                            if (innerSizes.Any())
                            {
                                int smallestInner = innerSizes.First();

                                foreach (var d in data)
                                {
                                    // Round down to nearest multiple of smallestInner
                                    d.Quantity = ((d.Quantity + smallestInner - 1) / smallestInner) * smallestInner;
                                }
                            }
                        }
                    }
                    catch
                    {
                        // Log or handle invalid JSON
                    }

                }

            }
            await _context.SaveChangesAsync();
            // Excel Report Path
            var reportPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "reports");
            Directory.CreateDirectory(reportPath);
            var filename = $"DuplicateTool {ProjectId}.xlsx";
            var filePath = Path.Combine(reportPath, filename);

            // Gather static properties (excluding NRDatas)
            var baseProperties = typeof(NRData).GetProperties(BindingFlags.Public | BindingFlags.Instance)
                                               .Where(p => p.Name != "NRDatas")
                                               .ToList();

            var extraHeaders = new HashSet<string>();
            var parsedRows = new List<(NRData row, Dictionary<string, string> extras)>();

            foreach (var row in reportRows)
            {
                var extras = new Dictionary<string, string>();

                if (!string.IsNullOrEmpty(row.NRDatas))
                {
                    try
                    {
                        extras = JsonSerializer.Deserialize<Dictionary<string, string>>(row.NRDatas);
                        if (extras != null)
                        {
                            foreach (var key in extras.Keys)
                                extraHeaders.Add(key);
                        }
                    }
                    catch
                    {
                        // Optionally log parsing error
                    }
                }

                parsedRows.Add((row, extras));
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

                var extraHeaderList = extraHeaders.OrderBy(k => k).ToList();
                foreach (var key in extraHeaderList)
                {
                    ws.Cells[1, col].Value = key;
                    ws.Cells[1, col].Style.Font.Bold = true;
                    col++;
                }

                // Write Rows
                int rowIdx = 2;
                foreach (var (item, extras) in parsedRows)
                {
                    col = 1;

                    foreach (var prop in baseProperties)
                    {
                        var value = prop.GetValue(item);
                        ws.Cells[rowIdx, col++].Value = value?.ToString() ?? "";
                    }

                    foreach (var key in extraHeaderList)
                    {
                        extras.TryGetValue(key, out var val);
                        ws.Cells[rowIdx, col++].Value = val ?? "";
                    }

                    // Highlight deleted rows
                    if (deletedRows.Any(x => x.Id == item.Id))
                    {
                        using (var range = ws.Cells[rowIdx, 1, rowIdx, col - 1])
                        {
                            range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                            range.Style.Fill.BackgroundColor.SetColor(Color.Red);
                        }
                    }

                    rowIdx++;
                }

                ws.Cells[ws.Dimension.Address].AutoFitColumns();
                package.SaveAs(new FileInfo(filePath));
            }

            return Ok(new
            {
                MergedRows = mergedCount,
                Consolidated = consolidate
            });

        }


        [HttpPost("EnvelopeConfiguration")]

        public async Task<IActionResult> EnvelopeConfiguration(int ProjectId)
        {
            var Envelopes = await _context.ProjectConfigs
                .Where(s => s.ProjectId == ProjectId)
                .Select(s => s.Envelope).FirstOrDefaultAsync();
            if(Envelopes == null)
            {
                return NotFound();
            }
            var NrData = await _context.NRDatas
                .Where(s => s.ProjectId == ProjectId)
               .ToListAsync();
            if (!NrData.Any())
            {
                return NotFound();
            }

            var envelopeDict = JsonSerializer.Deserialize<Dictionary<string, string>>(Envelopes);
            if (envelopeDict != null && envelopeDict.ContainsKey("Inner"))
            {
                var innerSizes = envelopeDict["Inner"]
                    .Split(',', StringSplitOptions.RemoveEmptyEntries)
                    .Select(e => e.Trim().ToUpper().Replace("E", "")) // get number from E10 → 10
                    .Where(x => int.TryParse(x, out _))
                    .Select(int.Parse)
                    .OrderByDescending(x => x)
                    .ToList();

                foreach (var row in NrData)
                {
                    int quantity = row.Quantity ?? 0;
                    var packetBreakdown = new Dictionary<string, string>(); // Size => Count

                    foreach (var size in innerSizes)
                    {
                        int count = quantity / size;
                        if (count > 0)
                        {
                            packetBreakdown[$"E{size}"] = count.ToString();
                            quantity -= count * size;
                        }
                    }
                    Dictionary<string, string> finalJson = new();

                    // Append or overwrite NRDatas column
                    // You can choose to overwrite or append — here's an example of appending
                    if (!string.IsNullOrWhiteSpace(row.NRDatas))
                    {
                        try
                        {
                            var existingData = JsonSerializer.Deserialize<Dictionary<string, JsonElement>>(row.NRDatas);
                            foreach (var kvp in existingData)
                            {
                                finalJson[kvp.Key] = kvp.Value.ToString(); // ensure string values
                            }
                        }
                        catch (JsonException)
                        {
                            // If existing NRDatas is invalid JSON, we ignore it
                        }
                    }
                    foreach (var kv in packetBreakdown)
                    {
                        finalJson[kv.Key] = kv.Value;
                    }

                    // Serialize final JSON back to string and save
                    row.NRDatas = JsonSerializer.Serialize(finalJson);
                }

                // Save to database

            }
            await _context.SaveChangesAsync();

            return Ok("Envelope breakdown successfully saved in NRDatas column.");
        }

    }
}