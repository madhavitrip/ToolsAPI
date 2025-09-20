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
using Newtonsoft.Json;
using Tools.Migrations;
using NRData = Tools.Models.NRData;

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
                    var subjectValues = group.Select(x => x.SubjectName?.Trim())
                                 .Where(v => !string.IsNullOrEmpty(v))
                                 .Distinct(StringComparer.OrdinalIgnoreCase)
                                 .ToList();

                    if (subjectValues.Count > 1)
                        keep.SubjectName = string.Join(" / ", subjectValues);
                    else if (subjectValues.Count == 1)
                        keep.SubjectName = subjectValues.First();

                    var courseValues = group.Select(x => x.CourseName?.Trim())
                                            .Where(v => !string.IsNullOrEmpty(v))
                                            .Distinct(StringComparer.OrdinalIgnoreCase)
                                            .ToList();

                    if (courseValues.Count > 1)
                        keep.CourseName = string.Join(" / ", courseValues);
                    else if (courseValues.Count == 1)
                        keep.CourseName = courseValues.First();
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
                        var envelopeDict = System.Text.Json.JsonSerializer.Deserialize<Dictionary<string, string>>(innerEnv);
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
                        extras = System.Text.Json.JsonSerializer.Deserialize<Dictionary<string, string>>(row.NRDatas);
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


        [HttpGet("MergeData")]
        public async Task<IActionResult> MergeEnvelope(int ProjectId)
        {
            // Helper function to parse JSON and sum the values (same as your existing one)
            int ParseJsonEnvelope(string value)
            {
                if (string.IsNullOrEmpty(value))
                {
                    return 0; // Return 0 if the value is null or empty
                }

                // Try parsing the string to an integer
                if (int.TryParse(value, out var result))
                {
                    return result;
                }
                else
                {
                    // Log invalid value if it cannot be parsed
                    Console.WriteLine($"Invalid value for envelope: {value}");
                    return 0;
                }
            }

            // Helper function to parse JSON and sum values for EnvelopeBreakages
            int ParseJsonEnvBreakage(string value)
            {
                if (string.IsNullOrEmpty(value))
                {
                    return 0; // Return 0 if the value is null or empty
                }

                // Sum values from JSON string like {\"E100\":\"1\",\"E50\":\"1\",\"E10\":\"3\"}
                try
                {
                    var envelopeData = Newtonsoft.Json.JsonConvert.DeserializeObject<Dictionary<string, int>>(value);
                    return envelopeData?.Values.Sum() ?? 0;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error parsing envelope data: {ex.Message}");
                    return 0;
                }
            }

            // Step 1: Group NRData by CatchNo and sum the quantity
            var NRdataGrouped = await _context.NRDatas
                .Where(p => p.ProjectId == ProjectId)
                .GroupBy(p => p.CatchNo)
                .Select(g => new
                {
                    CatchNo = g.Key,
                    TotalQuantity = g.Sum(p => p.Quantity),
                    NRDataId = g.Select(p => p.Id).ToList()
                })
                .ToListAsync();

            // Step 2: Fetch ExtrasEnvelope data and then group/sum in-memory
            var extraEnvelopeData = await _context.ExtrasEnvelope
                .Where(p => p.ProjectId == ProjectId)
                .ToListAsync();

            var ExtraEnvGrouped = extraEnvelopeData
                .GroupBy(p => p.CatchNo)
                .Select(g =>
                {
                    // Sum the quantities and the envelope values
                    var totalQuantity = g.Sum(p => p.Quantity);
                    var innerEnvelopeSum = g.Sum(p => ParseJsonEnvelope(p.InnerEnvelope));
                    var outerEnvelopeSum = g.Sum(p => ParseJsonEnvelope(p.OuterEnvelope));

                    // Log the summed values
                    Console.WriteLine($"CatchNo: {g.Key}, TotalQuantity: {totalQuantity}, InnerEnvelopeSum: {innerEnvelopeSum}, OuterEnvelopeSum: {outerEnvelopeSum}");

                    return new
                    {
                        CatchNo = g.Key,
                        TotalQuantity = totalQuantity,
                        InnerEnvelopeSum = innerEnvelopeSum,
                        OuterEnvelopeSum = outerEnvelopeSum
                    };
                })
                .ToList();

            // Step 3: Fetch EnvelopeBreakages data and group by NRDataIds, then sum in-memory
            var innerEnvData = await _context.EnvelopeBreakages
                .Where(p => p.ProjectId == ProjectId)
                .ToListAsync();

            var processedInnerEnv = innerEnvData
                .GroupBy(p => p.NrDataId) // Grouping by NrDataId
                .Select(g => new
                {
                    NRDataId = g.Key,
                    TotalInnerEnvelopes = g.Sum(p => ParseJsonEnvBreakage(p.InnerEnvelope)),
                    TotalOuterEnvelopes = g.Sum(p => ParseJsonEnvBreakage(p.OuterEnvelope))
                })
                .ToList();

            // Step 4: Merge NRDataGrouped and processedInnerEnv in-memory
            var mergedResult = from nr in NRdataGrouped
                               join ee in ExtraEnvGrouped on nr.CatchNo equals ee.CatchNo into extraEnvGroup
                               from extraEnv in extraEnvGroup.DefaultIfEmpty()
                               select new
                               {
                                   CatchNo = nr.CatchNo,
                                   NRDataIds = nr.NRDataId,
                                   TotalNRDataQuantity = nr.TotalQuantity + (extraEnv?.TotalQuantity ?? 0),
                                   TotalInnerEnvEnvelopes = processedInnerEnv
                               .Where(env => nr.NRDataId.Contains(env.NRDataId))
                               .Sum(env => env.TotalInnerEnvelopes)
                               + (extraEnv?.InnerEnvelopeSum ?? 0),  // Add ExtraEnv inner envelope

                                   TotalOuterEnvEnvelopes = processedInnerEnv
                               .Where(env => nr.NRDataId.Contains(env.NRDataId))
                               .Sum(env => env.TotalOuterEnvelopes)
                               + (extraEnv?.OuterEnvelopeSum ?? 0)
                               };

            return Ok(mergedResult);
        }



    }
}